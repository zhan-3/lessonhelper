"""Guarded, single-submit execution contract for HITWH course selection."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any
from urllib.parse import urlencode

from course_progress.collector import resolve_academic_url

from .selection_entry import extract_course_sections_from_html
from .selection_query import SELECTION_ENTRY_URL

_ALERT = re.compile(r"\balert\s*\(\s*(['\"])(.*?)\1\s*\)", re.DOTALL)


@dataclass(frozen=True)
class SelectionExecutionResult:
    status: str
    message: str
    section_id: str
    category: str
    term: str

    def to_dict(self) -> dict[str, str]:
        return {
            "status": self.status,
            "message": self.message,
            "section_id": self.section_id,
            "category": self.category,
            "term": self.term,
        }


def classify_selection_execution_html(html: str) -> tuple[str, str]:
    """Classify the server alert without retaining the returned HTML."""
    alerts = [" ".join(match.group(2).split()) for match in _ALERT.finditer(html)]
    message = next((item for item in alerts if item), "")
    if "选课成功" in message:
        return "selected", message
    if "总容量已满" in message:
        return "capacity_full", message
    if message:
        return "rejected", message
    return "unknown", "未识别到选课结果"


_READ_SECTION_PAGE = """
async ({queryUrl, category, termValue, pageNo}) => {
  const form = document.querySelector('form#queryform, form[name="queryform"]');
  if (!form) throw new Error('未找到选课表单');
  const parameters = new URLSearchParams(new FormData(form));
  parameters.set('rwh', '');
  parameters.set('pageXklb', category);
  parameters.set('pageXnxq', termValue);
  if (pageNo > 1) {
    parameters.set('pageNo', String(pageNo));
    parameters.set('pageSize', '20');
  } else {
    parameters.delete('pageNo');
    parameters.delete('pageSize');
  }
  const response = await fetch(queryUrl, {
    method: 'POST', credentials: 'same-origin', redirect: 'follow',
    headers: {'Content-Type': 'application/x-www-form-urlencoded'},
    body: parameters.toString(),
  });
  return {status: response.status, body: await response.text()};
}
"""


_EXECUTE_SELECTION = """
async ({submitUrl, sectionId, category, termValue}) => {
  const form = document.querySelector('form#queryform, form[name="queryform"]');
  if (!form) throw new Error('未找到选课表单');
  const parameters = new URLSearchParams(new FormData(form));
  parameters.set('rwh', sectionId);
  parameters.set('pageXklb', category);
  parameters.set('pageXnxq', termValue);
  const response = await fetch(submitUrl, {
    method: 'POST', credentials: 'same-origin', redirect: 'follow',
    headers: {'Content-Type': 'application/x-www-form-urlencoded'},
    body: parameters.toString(),
  });
  return {status: response.status, url: response.url, body: await response.text()};
}
"""


class VerifiedSelectionExecutionAdapter:
    """Reload one exact category and submit one exact page-provided rwh once."""

    def execute(
        self,
        page: Any,
        *,
        section_id: str,
        category: str,
        term_value: str,
        authenticate,
        source_page: int = 1,
    ) -> SelectionExecutionResult:
        if not section_id or not category or not term_value:
            raise ValueError("section_id, category and term are required")
        query_url = resolve_academic_url(page.url or SELECTION_ENTRY_URL, "/xsxk/queryXsxkList")
        entry = f"{query_url}?{urlencode({'pageXklb': category, 'pageXnxq': term_value})}"
        page = authenticate(entry, page)
        query = page.evaluate(
            _READ_SECTION_PAGE,
            {
                "queryUrl": query_url,
                "category": category,
                "termValue": term_value,
                "pageNo": max(1, int(source_page or 1)),
            },
        )
        if int(query.get("status") or 0) != 200:
            raise RuntimeError("重新查询教学班失败")
        matching = [
            item for item in extract_course_sections_from_html(str(query.get("body") or ""))
            if item.action_rwh == section_id
        ]
        if len(matching) != 1 or not matching[0].execution_ready:
            raise RuntimeError("教学班已不在当前待选课程中")
        submit_url = resolve_academic_url(page.url, "/xsxk/saveXsxk")
        response = page.evaluate(
            _EXECUTE_SELECTION,
            {
                "submitUrl": submit_url,
                "sectionId": section_id,
                "category": category,
                "termValue": term_value,
            },
        )
        status, message = classify_selection_execution_html(str(response.get("body") or ""))
        return SelectionExecutionResult(status, message, section_id, category, term_value)
