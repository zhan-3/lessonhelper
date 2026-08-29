"""Versioned, read-only adapter for the verified HITWH selection query."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import asdict
from types import SimpleNamespace
from typing import Any
from urllib.parse import parse_qs, urlencode

from playwright.sync_api import Error

from course_progress.academic_client import AuthenticatedAcademicClient
from course_progress.sanitizer import sanitize_request_body, sanitize_url

from .categories import CATEGORY_LABELS
from .selection_entry import (
    classify_selection_html,
    observation_to_dict,
    selection_page_count,
)

SELECTION_QUERY_CONTRACT_VERSION = "hitwh-jwts-selection-query-v1"
# Verified against the proxied jwts.hitwh.edu.cn student-selection page.  This
# URL establishes the page form and its dynamic token; subsequent reads use the
# form only as authenticated request context and never click a selection control.
SELECTION_ENTRY_URL = (
    "https://webvpn.hitwh.edu.cn/http/"
    "77726476706e69737468656265737421fae0558f693861446900c7a99c406d3667/"
    "xsxk/queryXsxk"
)


class SelectionContractError(RuntimeError):
    """The verified page or request contract is no longer available."""


def _sections_with_query_source(
    sections: dict[str, Any], category: str, *, semester: str = "", page: int | None = None
) -> dict[str, dict[str, Any]]:
    """Stamp each section with its query source.

    原始分类标签（如“专业基础课”）不能表达培养口径：外专业课程的标签
    与本院课完全同名。打上查询来源，前端据此映射培养要求筛选，而不
    是猜标签。
    """
    stamped: dict[str, dict[str, Any]] = {}
    for identity, section in sections.items():
        record = asdict(section)
        record["query_code"] = category
        record["query_label"] = CATEGORY_LABELS.get(category, category)
        record["query_term"] = semester
        record["query_page"] = page
        stamped[identity] = record
    return stamped


class VerifiedSelectionQueryAdapter:
    """Open one verified read page, then query every approved category directly."""

    version = SELECTION_QUERY_CONTRACT_VERSION

    def __init__(self, entry_url: str = SELECTION_ENTRY_URL):
        self.entry_url = entry_url

    def read(
        self,
        page: Any,
        *,
        categories: tuple[str, ...],
        semester_label: str,
        allowed_windows: dict[str, tuple[Any, ...]],
        progress: Callable[[str, dict[str, Any]], None],
        cancelled: Callable[[], bool],
        authenticate: Callable[[str, Any], Any] | None = None,
    ) -> dict[str, Any]:
        if not categories or not semester_label:
            return {"status": "no_matching_round", "sections": [], "queries": []}

        entry = f"{self.entry_url}?{urlencode({'pageXklb': categories[0]})}"
        auth = authenticate or (lambda url, target: (target.goto(url, wait_until="domcontentloaded", timeout=60_000), target)[1])
        client = AuthenticatedAcademicClient(page, authenticate=auth, timeout_seconds=15, cancelled=cancelled)
        try:
            client.get(entry)
            page = client.page
            frame = page.main_frame
            form = frame.locator('form#queryform, form[name="queryform"]').first
            if form.count() == 0:
                raise SelectionContractError("verified selection query form is missing")
        except SelectionContractError:
            raise
        except (Error, AttributeError, TypeError, ValueError) as error:
            raise SelectionContractError(
                f"verified selection entry is unavailable: {str(error)[:160]}"
            ) from error

        from course_progress.academic_client import resolve_academic_url

        endpoint = resolve_academic_url(frame.url, "/xsxk/queryXsxkList")
        queries: list[dict[str, Any]] = []
        all_sections: dict[str, dict[str, Any]] = {}
        for category in categories:
            pages: list[dict[str, Any]] = []
            sections: dict[str, Any] = {}
            stamped_sections: dict[str, dict[str, Any]] = {}
            expected_pages = 1
            last_result: dict[str, Any] = {}
            try:
                for page_number in range(1, 51):
                    if page_number > expected_pages or cancelled():
                        break
                    progress(
                        "reading",
                        {
                            "target": "selection",
                            "contract_version": self.version,
                            "category": category,
                            "page": page_number,
                            "page_count": expected_pages,
                        },
                    )
                    overrides = {"rwh": "", "pageXklb": category}
                    remove: tuple[str, ...] = ()
                    if page_number > 1:
                        overrides.update({"pageNo": str(page_number), "pageSize": "20", "pageCount": str(expected_pages)})
                    else:
                        remove = ("pageNo", "pageSize")
                    response = client.post_page_form(
                        "/xsxk/queryXsxkList", overrides=overrides, remove=remove,
                    )
                    last_result = {
                        "status": response.status, "url": response.url,
                        "requestBody": response.request_body, "body": response.body,
                    }
                    html = response.body
                    if page_number == 1:
                        expected_pages = min(selection_page_count(html), 50)
                    expected = tuple(
                        item if not isinstance(item, dict) else SimpleNamespace(**item)
                        for item in allowed_windows.get(category, ())
                    )
                    observation = classify_selection_html(
                        int(last_result["status"]),
                        html,
                        request_url=str(last_result.get("url") or endpoint),
                        expected_windows=expected,
                    )
                    request_values = parse_qs(str(last_result.get("requestBody") or ""))
                    term_value = str((request_values.get("pageXnxq") or [""])[0])
                    for section in observation.sections:
                        sections.setdefault(section.identity, section)
                        stamped_sections.setdefault(
                            section.identity,
                            _sections_with_query_source(
                                {section.identity: section}, category,
                                semester=term_value, page=page_number,
                            )[section.identity],
                        )
                    pages.append(
                        {
                            "page": page_number,
                            "observation": observation_to_dict(observation),
                        }
                    )
                complete = not cancelled() and len(pages) == expected_pages
                query = {
                    "category": category,
                    "semester": semester_label,
                    "contract_version": self.version,
                    "method": "POST",
                    "url": sanitize_url(str(last_result.get("url") or endpoint)),
                    "request_body": sanitize_request_body(
                        str(last_result.get("requestBody") or "")
                    ),
                    "page_count": expected_pages,
                    "pages_fetched": len(pages),
                    "complete": complete,
                    "record_count": len(sections),
                    "sections": list(stamped_sections.values()),
                    "pages": pages,
                }
                queries.append(query)
                for identity, record in stamped_sections.items():
                    all_sections.setdefault(identity, record)
            except (Error, KeyError, TypeError, ValueError) as error:
                # Keep a category-level incomplete result.  The caller will not
                # publish it as a current snapshot, and a transient page error
                # must not trigger a second discovery/query pass.
                queries.append(
                    {
                        "category": category,
                        "semester": semester_label,
                        "contract_version": self.version,
                        "complete": False,
                        "page_count": expected_pages,
                        "pages_fetched": len(pages),
                        "record_count": len(sections),
                        "sections": list(stamped_sections.values()),
                        "pages": pages,
                        "error": str(error)[:160],
                    }
                )

        complete = bool(queries) and all(item.get("complete") for item in queries)
        return {
            "status": "complete" if complete else "incomplete",
            "contract_version": self.version,
            "sections": list(all_sections.values()),
            "queries": queries,
            "_trace_requests": client.trace_requests,
        }
