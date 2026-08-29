"""Parser for the verified read-only personal timetable response."""

from __future__ import annotations

import re
from html.parser import HTMLParser

from .timetable import TimetableEntry


FETCH_TIMETABLE_PAGE_SCRIPT = """
async ({url}) => {
  const form = document.querySelector('form#queryform, form[name="queryform"]');
  if (!form) throw new Error('未找到课表查询表单');
  const parameters = new URLSearchParams(new FormData(form));
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


_ALIASES = {
    "course_code": ("课程代码", "课程编号", "课程号", "course_code"),
    "course_name": ("课程名称", "课程名", "course_name"),
    "teaching_class": ("教学班号", "教学班", "班号", "teaching_class"),
    "teacher": ("教师", "任课教师", "教师姓名", "teacher"),
    "weekday": ("星期", "上课星期", "weekday"),
    "periods": ("节次", "上课节次", "上课时间", "periods"),
    "weeks": ("周次", "教学周", "weeks"),
    "week_parity": ("单双周", "周次类型", "week_parity"),
    "location": ("上课地点", "上课教室", "教室", "地点", "location"),
}


def personal_timetable_parameters(*, fhlj: str, xnxq: str) -> dict[str, str]:
    """Build the exact verified read-only request body."""
    values = {"fhlj": _clean(fhlj), "xnxq": _clean(xnxq)}
    if not all(values.values()):
        raise ValueError("personal timetable request requires fhlj and xnxq")
    return values


def _clean(value: str) -> str:
    return " ".join(value.split()).strip()


def _normalized(value: str) -> str:
    return re.sub(r"[：:：\s]", "", _clean(value)).lower()


def _clean_cell(value: str) -> str:
    """Normalize cell text while preserving HTML line-break boundaries."""
    return "\n".join(part for line in value.splitlines() if (part := _clean(line)))


class _TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.tables: list[list[list[str]]] = []
        self._table: list[list[str]] | None = None
        self._row: list[str] | None = None
        self._cell: list[str] | None = None
        self._in_cell = False

    def handle_starttag(self, tag: str, attrs) -> None:
        tag = tag.lower()
        if tag == "table" and self._table is None:
            self._table = []
        elif tag == "tr" and self._table is not None:
            self._row = []
        elif tag in {"th", "td"} and self._row is not None:
            self._cell = []
            self._in_cell = True
        elif tag == "br" and self._in_cell and self._cell is not None:
            self._cell.append("\n")

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        # 教务系统用非法的 ``</br>`` 结束标签充当换行符(而非 ``<br>``)。
        if tag == "br" and self._in_cell and self._cell is not None:
            self._cell.append("\n")
        elif tag in {"th", "td"} and self._in_cell and self._row is not None:
            self._row.append(_clean_cell("".join(self._cell or [])))
            self._cell = None
            self._in_cell = False
        elif tag == "tr" and self._table is not None and self._row is not None:
            if any(self._row):
                self._table.append(self._row)
            self._row = None
        elif tag == "table" and self._table is not None:
            self.tables.append(self._table)
            self._table = None

    def handle_data(self, data: str) -> None:
        if self._in_cell and self._cell is not None:
            self._cell.append(data)


def _parse_range(value: str, label: str) -> tuple[int, int]:
    numbers = [int(item) for item in re.findall(r"\d+", value)]
    if not numbers:
        raise ValueError(f"unable to parse {label}: {value!r}")
    if len(numbers) == 1:
        return numbers[0], numbers[0]
    return min(numbers[0], numbers[1]), max(numbers[0], numbers[1])


def _parse_weeks(value: str) -> tuple[int, ...]:
    if re.search(r"待定|未定|pending|tbd", value, re.IGNORECASE):
        return ()
    result: set[int] = set()
    for match in re.finditer(r"(\d+)\s*[-~至]\s*(\d+)|(\d+)", value):
        if match.group(1):
            result.update(range(min(int(match.group(1)), int(match.group(2))), max(int(match.group(1)), int(match.group(2))) + 1))
        else:
            result.add(int(match.group(3)))
    return tuple(sorted(result))


def _parse_weekday(value: str) -> int:
    chinese = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "日": 7, "天": 7}
    match = re.search(r"[1-7]", value)
    if match:
        return int(match.group())
    for char, number in chinese.items():
        if char in value:
            return number
    raise ValueError(f"unable to parse weekday: {value!r}")


def _parity(value: str) -> str:
    if "单" in value:
        return "odd"
    if "双" in value:
        return "even"
    return "all"


def _find_table(html: str) -> tuple[list[list[str]], dict[str, int]]:
    parser = _TableParser()
    parser.feed(html)
    for table in parser.tables:
        for index, row in enumerate(table):
            headers = {_normalized(cell): column for column, cell in enumerate(row)}
            mapped: dict[str, int] = {}
            for field, aliases in _ALIASES.items():
                for alias in aliases:
                    if _normalized(alias) in headers:
                        mapped[field] = headers[_normalized(alias)]
                        break
            if {"course_name", "weekday", "periods", "weeks"} <= mapped.keys():
                return table[index + 1 :], mapped
    raise ValueError("required timetable columns were not found")


def parse_personal_timetable_html(html: str, *, term: str) -> tuple[TimetableEntry, ...]:
    """Parse a complete personal-timetable table without retaining HTML."""
    rows, headers = _find_table(html)
    entries: list[TimetableEntry] = []
    for row in rows:
        def value(field: str) -> str:
            index = headers.get(field)
            return _clean(row[index]) if index is not None and index < len(row) else ""

        course_name = value("course_name")
        if not course_name:
            continue
        weeks = _parse_weeks(value("weeks"))
        periods = value("periods")
        try:
            weekday = _parse_weekday(value("weekday"))
        except ValueError:
            weekday = None
        try:
            start_period, end_period = _parse_range(periods, "periods")
        except ValueError:
            start_period = end_period = None
        unknown = (
            weekday is None
            or not weeks
            or start_period is None
            or bool(re.search(r"待定|未定|pending|tbd", periods, re.IGNORECASE))
        )
        if unknown:
            start_period = end_period = None
        entries.append(
            TimetableEntry(
                term=term,
                course_code=value("course_code"),
                course_name=course_name,
                weekday=weekday,
                start_period=start_period,
                end_period=end_period,
                week_start=min(weeks) if weeks else None,
                week_end=max(weeks) if weeks else None,
                week_parity=_parity(value("week_parity")),
                teaching_class=value("teaching_class"),
                teacher=value("teacher"),
                location=value("location"),
                week_numbers=weeks,
                conflict_status="unknown" if unknown else "known",
            )
        )
    if not entries:
        raise ValueError("personal timetable contained no usable course records")
    return tuple(entries)


def _extract_grid_term(html: str) -> str:
    """Extract the 「20xx秋季学期」 title from a grid timetable page."""
    match = re.search(r"(20\d{2})\s*(春季|夏季|秋季|冬季)学期", html)
    if not match:
        return ""
    return _clean(f"{match.group(1)}{match.group(2)}学期")


def _parse_other_course_row(cell: str, term: str) -> list[TimetableEntry]:
    """Parse a single 「其它课程： 课程名◇教师◇周次◇…」 column-span row."""
    text = re.sub(r"^其它课程[:：]\s*", "", _clean(cell))
    if not text:
        return []
    parts = [part.strip() for part in text.split("◇") if part.strip()]
    if not parts:
        return []
    course_name = parts[0]
    week_text = ""
    teacher = ""
    for part in parts[1:]:
        if not week_text and re.fullmatch(r"[\d,，、\s\-~至]+", part):
            week_text = part.replace("周", "")
        elif not teacher and not re.search(r"\d", part):
            teacher = part
    week_numbers = _parse_weeks(week_text) if week_text else ()
    return [
        TimetableEntry(
            term=term,
            course_code="",
            course_name=course_name,
            weekday=None,
            start_period=None,
            end_period=None,
            week_start=min(week_numbers) if week_numbers else None,
            week_end=max(week_numbers) if week_numbers else None,
            week_parity="all",
            teacher=teacher,
            location="",
            week_numbers=week_numbers,
            conflict_status="unknown",
        )
    ]


def parse_timetable_grid_html(
    html: str, *, expected_term: str | None = None
) -> tuple[TimetableEntry, ...]:
    """Parse the verified 星期×节次 grid returned by POST /kbcx/queryGrkb."""
    from .timetable import import_grid_timetable_rows, normalize_timetable_term

    parser = _TableParser()
    parser.feed(html)
    grid = None
    for table in parser.tables:
        if any(_clean(cell).startswith("星期") for row in table for cell in row):
            grid = table
            break
    if grid is None:
        raise ValueError("课表页面未找到星期网格表格")

    header_index = next(
        index
        for index, row in enumerate(grid)
        if any(_clean(cell).startswith("星期") for cell in row)
    )
    header_row = grid[header_index]
    body_rows = grid[header_index + 1 :]

    # 其它课程行是 <td colspan=9> 单列行;节次行至少两列(时段/节次)。
    other_rows = [row for row in body_rows if len(row) < 2]
    period_rows = [row for row in body_rows if len(row) >= 2]

    # 星期头是 <th colspan=2> 合并首两列,比节次行少一列;补空列对齐。
    aligned_header = list(header_row)
    period_width = max((len(row) for row in period_rows), default=len(aligned_header))
    if period_width > len(aligned_header):
        aligned_header.insert(1, "")

    term_text = _extract_grid_term(html) or (expected_term or "")
    rows = [[term_text], aligned_header, *period_rows]
    entries = list(import_grid_timetable_rows(rows, expected_term)) if period_rows else []

    normalized_term = normalize_timetable_term(term_text) if term_text else ""
    for row in other_rows:
        if row:
            entries.extend(_parse_other_course_row(row[0], normalized_term))

    if not entries:
        raise ValueError("课表页面没有可用课程记录")
    return tuple(entries)
