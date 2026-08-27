"""Versioned, read-only adapter for the verified HITWH selection query."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import asdict
from types import SimpleNamespace
from typing import Any
from urllib.parse import urlencode

from playwright.sync_api import Error

from course_progress.sanitizer import sanitize_request_body, sanitize_url

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

_FETCH_SELECTION_PAGE = """
async ({url, category, semesterLabel, pageNo, pageCount}) => {
  const form = document.querySelector('form#queryform, form[name="queryform"]');
  if (!form) throw new Error('未找到选课查询表单');
  const parameters = new URLSearchParams(new FormData(form));
  parameters.set('rwh', '');
  parameters.set('pageXklb', category);
  if (pageNo !== null) {
    parameters.set('pageNo', String(pageNo));
    parameters.set('pageSize', '20');
    parameters.set('pageCount', String(pageCount));
  } else {
    parameters.delete('pageNo');
    parameters.delete('pageSize');
    parameters.delete('pageCount');
  }
  const semester = Array.from(form.querySelectorAll('select[name="pageXnxq"] option'))
    .find(option => option.textContent.replace(/[\\s年]|学期/g, '') === semesterLabel);
  if (!semester) throw new Error(`通知学期不在页面选项中: ${semesterLabel}`);
  parameters.set('pageXnxq', semester.value);
  const response = await fetch(url, {
    method: 'POST',
    credentials: 'same-origin',
    headers: {'Content-Type': 'application/x-www-form-urlencoded'},
    body: parameters.toString(),
    redirect: 'follow',
  });
  return {
    status: response.status,
    url: response.url,
    requestBody: parameters.toString(),
    body: await response.text(),
  };
}
"""


class SelectionContractError(RuntimeError):
    """The verified page or request contract is no longer available."""


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
        try:
            if authenticate is None:
                page.goto(entry, wait_until="domcontentloaded", timeout=60_000)
            else:
                page = authenticate(entry, page)
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

        from course_progress.collector import resolve_academic_url

        endpoint = resolve_academic_url(frame.url, "/xsxk/queryXsxkList")
        queries: list[dict[str, Any]] = []
        all_sections: dict[str, dict[str, Any]] = {}
        for category in categories:
            pages: list[dict[str, Any]] = []
            sections: dict[str, Any] = {}
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
                    last_result = frame.evaluate(
                        _FETCH_SELECTION_PAGE,
                        {
                            "url": endpoint,
                            "category": category,
                            "semesterLabel": semester_label,
                            "pageNo": None if page_number == 1 else page_number,
                            "pageCount": expected_pages,
                        },
                    )
                    html = str(last_result["body"])
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
                    for section in observation.sections:
                        sections.setdefault(section.identity, section)
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
                    "sections": [asdict(section) for section in sections.values()],
                    "pages": pages,
                }
                queries.append(query)
                for identity, section in sections.items():
                    all_sections.setdefault(identity, asdict(section))
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
                        "sections": [asdict(section) for section in sections.values()],
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
        }
