"""Collect passed-course records across dynamic semesters and result pages."""

from __future__ import annotations

import json
import re
from collections import Counter
from dataclasses import asdict, dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Callable, Iterable

from .academic_client import (
    AcademicContractError,
    AuthenticatedAcademicClient,
)
from .progress import AcademicRecord, parse_grade_html


GRADE_ENDPOINT = "/cjcx/queryQmcj"


def grade_query_parameters(
    semester: str, page_number: int, *, page_size: int = 20
) -> dict[str, str]:
    parameters = {
        "pageXnxq": semester,
        "pageBkcxbj": "",
        # The portal's server-side passed filter can return an empty table for
        # a valid semester. Fetch the semester's records and filter locally
        # from the final-grade field instead.
        "pageSfjg": "",
        "pageKcmc": "",
    }
    # The portal uses a different form for the initial filtered query. Its
    # first-page POST contains only the four filters above; pagination fields
    # are added only when navigating to page 2 and beyond.
    if page_number > 1:
        parameters.update(
            {
                "pageNo": str(page_number),
                "pageSize": str(page_size),
            }
        )
    return parameters


@dataclass(frozen=True)
class SemesterOption:
    value: str
    label: str


def format_semester_label(value: str, label: str) -> str:
    """Turn portal semester codes into compact, human-readable labels."""
    match = re.fullmatch(r"(\d{4})-(\d{4})([123])", value.strip())
    if match:
        start_year, end_year, term = match.groups()
        year = start_year if term == "1" else end_year
        season = {"1": "秋季", "2": "春季", "3": "夏季"}[term]
        return f"{year}{season}"
    return label.strip() or value.strip()


def format_collection_summary(
    semesters: Iterable[SemesterOption], records: Iterable[AcademicRecord]
) -> str:
    """Summarize parsed course counts without printing request-level details."""
    counts = Counter(record.semester for record in records)
    return "；".join(
        f"{format_semester_label(item.value, item.label)} {counts[item.value]} 条"
        for item in semesters
    )


@dataclass(frozen=True)
class CollectionFailure:
    semester_value: str
    semester_label: str
    page_number: int
    message: str


class SessionExpiredError(ValueError):
    """The academic endpoint returned an authentication page."""


@dataclass(frozen=True)
class CollectionCheckpoint:
    records: tuple[AcademicRecord, ...] = ()
    completed_pages: tuple[tuple[str, int], ...] = ()
    page_counts: tuple[tuple[str, int], ...] = ()


def save_checkpoint(path: Path, checkpoint: CollectionCheckpoint) -> None:
    """Persist only resumable completion facts, never raw responses or scores."""
    payload = {
        "records": [asdict(record) for record in checkpoint.records],
        "completed_pages": [
            {"semester": semester, "page": page}
            for semester, page in checkpoint.completed_pages
        ],
        "page_counts": {
            semester: page_count for semester, page_count in checkpoint.page_counts
        },
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), "utf-8")
    temporary.replace(path)


def load_checkpoint(path: Path) -> CollectionCheckpoint:
    """Load a checkpoint, treating an absent file as an empty collection."""
    if not path.is_file():
        return CollectionCheckpoint()
    payload = json.loads(path.read_text("utf-8"))
    records = tuple(
        AcademicRecord(
            semester=str(item["semester"]),
            code=str(item["code"]),
            name=str(item["name"]),
            nature=str(item["nature"]),
            category=str(item["category"]),
            credits=float(item["credits"]),
            passed=bool(item["passed"]),
        )
        for item in payload.get("records", [])
    )
    completed_pages = tuple(
        (str(item["semester"]), int(item["page"]))
        for item in payload.get("completed_pages", [])
    )
    page_counts = tuple(
        (str(semester), int(page_count))
        for semester, page_count in payload.get("page_counts", {}).items()
    )
    return CollectionCheckpoint(records, completed_pages, page_counts)


@dataclass(frozen=True)
class GradeCollection:
    records: tuple[AcademicRecord, ...]
    semesters: tuple[SemesterOption, ...]
    failures: tuple[CollectionFailure, ...]

    @property
    def complete(self) -> bool:
        return not self.failures


class _PageCountParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.page_count = 1

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag != "input":
            return
        attributes = dict(attrs)
        if attributes.get("id") != "pageCount":
            return
        try:
            self.page_count = max(1, int(attributes.get("value", "1")))
        except ValueError:
            self.page_count = 1


def parse_page_count(html: str) -> int:
    parser = _PageCountParser()
    parser.feed(html)
    return parser.page_count


class _SemesterParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.in_selector = False
        self.current_value = ""
        self.current_text: list[str] = []
        self.options: list[SemesterOption] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        attributes = dict(attrs)
        if tag == "select" and attributes.get("id") == "xnxqid":
            self.in_selector = True
        elif tag == "option" and self.in_selector:
            self.current_value = str(attributes.get("value") or "").strip()
            self.current_text = []

    def handle_data(self, data: str) -> None:
        if self.in_selector and self.current_value:
            self.current_text.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag == "option" and self.in_selector and self.current_value:
            self.options.append(SemesterOption(self.current_value, "".join(self.current_text).strip()))
            self.current_value = ""
            self.current_text = []
        elif tag == "select" and self.in_selector:
            self.in_selector = False


def parse_semester_options(html: str) -> tuple[SemesterOption, ...]:
    parser = _SemesterParser()
    parser.feed(html)
    if not parser.options:
        raise AcademicContractError("成绩接口缺少学期选择器 #xnxqid")
    return tuple(parser.options)


def _validate_grade_page(html: str) -> None:
    lowered = html.lower()
    if "统一身份认证" in html or "authserver/login" in lowered:
        raise SessionExpiredError("响应中检测到统一身份认证页面，登录可能已失效")
    if "课程代码" not in html or "最终成绩" not in html:
        raise ValueError("响应中未找到成绩表，登录可能已失效或页面结构已变化")


def collect_grade_records(
    semesters: Iterable[SemesterOption],
    fetch_page: Callable[[str, int], str],
    *,
    on_page: Callable[[SemesterOption, int, int], None] | None = None,
    on_page_data: Callable[
        [SemesterOption, int, int, tuple[AcademicRecord, ...]], None
    ]
    | None = None,
    on_session_expired: Callable[[], None] | None = None,
    checkpoint: CollectionCheckpoint | None = None,
    is_cancelled: Callable[[], bool] | None = None,
) -> GradeCollection:
    """Collect all result pages; preserve failures instead of treating them as zero."""
    semester_list = tuple(semesters)
    checkpoint = checkpoint or CollectionCheckpoint()
    records: list[AcademicRecord] = list(checkpoint.records)
    failures: list[CollectionFailure] = []
    completed_pages = set(checkpoint.completed_pages)
    page_counts = dict(checkpoint.page_counts)

    def fetch_validated_page(semester: SemesterOption, page_number: int) -> str:
        while True:
            try:
                html = fetch_page(semester.value, page_number)
                _validate_grade_page(html)
                return html
            except SessionExpiredError:
                if on_session_expired is None:
                    raise
                on_session_expired()

    for semester in semester_list:
        if is_cancelled is not None and is_cancelled():
            failures.append(CollectionFailure(semester.value, semester.label, 0, "cancelled"))
            break
        if (semester.value, 1) in completed_pages and semester.value in page_counts:
            page_count = page_counts[semester.value]
        else:
            try:
                first_html = fetch_validated_page(semester, 1)
                first_records = parse_grade_html(first_html)
                records.extend(first_records)
                page_count = parse_page_count(first_html)
                page_counts[semester.value] = page_count
                completed_pages.add((semester.value, 1))
                if on_page is not None:
                    on_page(semester, 1, len(first_records))
                if on_page_data is not None:
                    on_page_data(semester, 1, page_count, tuple(first_records))
            except Exception as exc:
                failures.append(
                    CollectionFailure(
                        semester.value, semester.label, 1, str(exc) or type(exc).__name__
                    )
                )
                continue

        for page_number in range(2, page_count + 1):
            if is_cancelled is not None and is_cancelled():
                failures.append(CollectionFailure(semester.value, semester.label, page_number, "cancelled"))
                return GradeCollection(tuple(records), semester_list, tuple(failures))
            if (semester.value, page_number) in completed_pages:
                continue
            try:
                page_html = fetch_validated_page(semester, page_number)
                page_records = parse_grade_html(page_html)
                records.extend(page_records)
                completed_pages.add((semester.value, page_number))
                if on_page is not None:
                    on_page(semester, page_number, len(page_records))
                if on_page_data is not None:
                    on_page_data(
                        semester, page_number, page_count, tuple(page_records)
                    )
            except Exception as exc:
                failures.append(
                    CollectionFailure(
                        semester.value,
                        semester.label,
                        page_number,
                        str(exc) or type(exc).__name__,
                    )
                )

    return GradeCollection(tuple(records), semester_list, tuple(failures))


class FixedGradeReader:
    """Read every grade page through the fixed authenticated request contract."""

    def __init__(self, client: AuthenticatedAcademicClient, *, page_size: int = 20):
        self.client = client
        self.page_size = page_size

    def collect(
        self,
        *,
        on_page_data: Callable[[SemesterOption, int, int, tuple[AcademicRecord, ...]], None] | None = None,
        is_cancelled: Callable[[], bool] | None = None,
        checkpoint: CollectionCheckpoint | None = None,
    ) -> GradeCollection:
        entry = self.client.get(GRADE_ENDPOINT)
        semesters = parse_semester_options(entry.body)

        def fetch_page(semester: str, page_number: int) -> str:
            response = self.client.post_form(
                GRADE_ENDPOINT,
                grade_query_parameters(semester, page_number, page_size=self.page_size),
                retry_read_once=True,
            )
            if response.status != 200:
                raise RuntimeError(f"成绩查询失败: HTTP {response.status}")
            return response.body

        return collect_grade_records(
            semesters,
            fetch_page,
            on_page_data=on_page_data,
            is_cancelled=is_cancelled,
            checkpoint=checkpoint,
        )

