"""Parse the verified current-term 「学生选课信息查询」 response."""

from __future__ import annotations

from dataclasses import dataclass
from html.parser import HTMLParser


@dataclass(frozen=True)
class EnrolledCourse:
    term: str
    code: str
    name: str
    category: str
    nature: str
    credits: float


class _TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tables: list[list[list[str]]] = []
        self._table: list[list[str]] | None = None
        self._row: list[str] | None = None
        self._cell: list[str] | None = None

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag == "table" and self._table is None:
            self._table = []
        elif tag == "tr" and self._table is not None:
            self._row = []
        elif tag in {"th", "td"} and self._row is not None:
            self._cell = []

    def handle_data(self, data: str) -> None:
        if self._cell is not None:
            self._cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag in {"th", "td"} and self._row is not None and self._cell is not None:
            self._row.append(" ".join("".join(self._cell).split()))
            self._cell = None
        elif tag == "tr" and self._table is not None and self._row is not None:
            if self._row:
                self._table.append(self._row)
            self._row = None
        elif tag == "table" and self._table is not None:
            if self._table:
                self.tables.append(self._table)
            self._table = None


def parse_current_enrollment_html(html: str) -> tuple[EnrolledCourse, ...]:
    """Return only score-free identity/category/credit facts from the fixed table."""
    parser = _TableParser()
    parser.feed(html)
    required = ("学年学期", "课程代码", "课程名称", "课程类别", "课程性质", "学分")
    table: list[list[str]] | None = None
    headers: list[str] = []
    for candidate in parser.tables:
        for index, row in enumerate(candidate):
            if all(field in row for field in required):
                headers = row
                table = candidate[index + 1 :]
                break
        if table is not None:
            break
    if table is None:
        raise ValueError("已选课程表缺少必要字段")

    positions = {field: headers.index(field) for field in required}
    courses: dict[str, EnrolledCourse] = {}
    for row in table:
        if len(row) <= max(positions.values()):
            continue
        name = row[positions["课程名称"]].strip()
        code = row[positions["课程代码"]].strip()
        if not name:
            continue
        try:
            credits = float(row[positions["学分"]])
        except ValueError:
            continue
        identity = code or " ".join(name.lower().split())
        courses.setdefault(
            identity,
            EnrolledCourse(
                term=row[positions["学年学期"]].strip(),
                code=code,
                name=name,
                category=row[positions["课程类别"]].strip(),
                nature=row[positions["课程性质"]].strip(),
                credits=credits,
            ),
        )
    return tuple(courses.values())
