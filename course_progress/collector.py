"""Collect passed-course records across dynamic semesters and result pages."""

from __future__ import annotations

import re
import time
from dataclasses import dataclass
from html.parser import HTMLParser
from typing import Callable, Iterable
from urllib.parse import urljoin, urlsplit

from playwright.sync_api import Frame, Page

from .progress import AcademicRecord, parse_grade_html


GRADE_ENDPOINT = "/cjcx/queryQmcj"


def resolve_academic_url(current_url: str, endpoint: str) -> str:
    """Resolve an academic endpoint for direct access or a WebVPN rewritten app."""
    parts = urlsplit(current_url)
    match = re.match(r"^(/http/[0-9a-fA-F]+)(?:/|$)", parts.path)
    if match:
        return f"{parts.scheme}://{parts.netloc}{match.group(1)}/{endpoint.lstrip('/')}"
    return urljoin(current_url, endpoint)


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


@dataclass(frozen=True)
class CollectionFailure:
    semester_value: str
    semester_label: str
    page_number: int
    message: str


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


def _validate_grade_page(html: str) -> None:
    if "课程代码" not in html or "最终成绩" not in html:
        raise ValueError("响应中未找到成绩表，登录可能已失效或页面结构已变化")


def collect_grade_records(
    semesters: Iterable[SemesterOption],
    fetch_page: Callable[[str, int], str],
    *,
    on_page: Callable[[SemesterOption, int, int], None] | None = None,
) -> GradeCollection:
    """Collect all result pages; preserve failures instead of treating them as zero."""
    semester_list = tuple(semesters)
    records: list[AcademicRecord] = []
    failures: list[CollectionFailure] = []

    for semester in semester_list:
        try:
            first_html = fetch_page(semester.value, 1)
            _validate_grade_page(first_html)
            first_records = parse_grade_html(first_html)
            records.extend(first_records)
            if on_page is not None:
                on_page(semester, 1, len(first_records))
            page_count = parse_page_count(first_html)
        except Exception as exc:
            failures.append(
                CollectionFailure(
                    semester.value, semester.label, 1, str(exc) or type(exc).__name__
                )
            )
            continue

        for page_number in range(2, page_count + 1):
            try:
                page_html = fetch_page(semester.value, page_number)
                _validate_grade_page(page_html)
                page_records = parse_grade_html(page_html)
                records.extend(page_records)
                if on_page is not None:
                    on_page(semester, page_number, len(page_records))
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


_FETCH_GRADE_PAGE = """
async ({url, parameters}) => {
  const body = new URLSearchParams(parameters).toString();
  const response = await fetch(url, {
    method: 'POST',
    credentials: 'same-origin',
    headers: {'Content-Type': 'application/x-www-form-urlencoded'},
    body,
    redirect: 'follow',
  });
  if (!response.ok) {
    throw new Error(`成绩查询失败: HTTP ${response.status}`);
  }
  return {
    status: response.status,
    length: Number(response.headers.get('content-length') || 0),
    body: await response.text(),
  };
}
"""


def find_academic_frame(pages: Iterable[Page]) -> Frame | None:
    for page in pages:
        frame = page.frame(name="iframename")
        if frame is not None and frame.url != "about:blank":
            return frame
    return None


def wait_for_academic_frame(page: Page, timeout_seconds: int = 300) -> Frame:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        pages = tuple(page.context.pages) or (page,)
        frame = find_academic_frame(pages)
        if frame is not None:
            return frame
        page.wait_for_timeout(500)
    raise TimeoutError("等待教务系统主内容 iframe 超时")


class PlaywrightGradeCollector:
    """Read passed grade pages through an authenticated academic-system frame."""

    def __init__(self, *, page_size: int = 20):
        if page_size <= 0:
            raise ValueError("page_size 必须大于 0")
        self.page_size = page_size

    def collect(
        self, page: Page, *, frame_timeout_seconds: int = 300
    ) -> GradeCollection:
        print("正在等待教务系统标签页和课程 iframe……")
        frame = wait_for_academic_frame(page, frame_timeout_seconds)
        print("已找到教务系统课程 iframe，正在读取学期选项……")
        grade_url = resolve_academic_url(frame.url, GRADE_ENDPOINT)
        frame.goto(grade_url, wait_until="domcontentloaded", timeout=60_000)

        options = frame.locator("#xnxqid option").evaluate_all(
            "options => options.map(option => ({value: option.value, label: option.textContent.trim()}))"
        )
        semesters = tuple(
            SemesterOption(str(option["value"]), str(option["label"]))
            for option in options
            if str(option["value"]).strip()
        )
        if not semesters:
            raise RuntimeError("期末成绩页面未提供任何学期选项")
        print(f"已读取 {len(semesters)} 个学期，开始逐学期采集成绩记录……")

        def fetch_page(semester: str, page_number: int) -> str:
            result = frame.evaluate(
                _FETCH_GRADE_PAGE,
                {
                    "url": grade_url,
                    "parameters": grade_query_parameters(
                        semester, page_number, page_size=self.page_size
                    ),
                },
            )
            body = str(result["body"])
            print(
                f"接口响应：学期 {semester} 第 {page_number} 页，"
                f"HTTP {result['status']}，响应 {len(body)} 字符"
            )
            return body

        def report_page(
            semester: SemesterOption, page_number: int, record_count: int
        ) -> None:
            print(
                f"接口解析：{semester.label} 第 {page_number} 页，"
                f"课程记录 {record_count} 条"
            )

        return collect_grade_records(semesters, fetch_page, on_page=report_page)
