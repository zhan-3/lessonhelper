"""Calculate course progress from raw passed-course records and guidance Markdown."""

from __future__ import annotations

import re
from collections import defaultdict
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path


@dataclass(frozen=True)
class CompletedCourse:
    code: str
    name: str
    nature: str
    category: str
    credits: float


@dataclass(frozen=True)
class AcademicRecord:
    semester: str
    code: str
    name: str
    nature: str
    category: str
    credits: float
    passed: bool
    # A blank final-grade cell is a current, selected course rather than a
    # failed course. Keep the distinction without retaining the grade itself.
    in_progress: bool = False


@dataclass(frozen=True)
class Requirement:
    key: str
    label: str
    minimum_credits: float
    contribution_keys: tuple[str, ...] = ()


@dataclass(frozen=True)
class RequirementBaseline:
    version: str
    requirements: tuple[Requirement, ...]
    category_mapping: Mapping[str, str]


@dataclass(frozen=True)
class Progress:
    requirement: Requirement
    completed_credits: float
    courses: tuple[CompletedCourse, ...]
    in_progress_courses: tuple[CompletedCourse, ...] = ()

    @property
    def remaining_credits(self) -> float:
        return max(0.0, self.requirement.minimum_credits - self.completed_credits)


@dataclass(frozen=True)
class CourseConflict:
    identity: str
    records: tuple[AcademicRecord, ...]


@dataclass(frozen=True)
class ProgressReport:
    baseline_version: str
    progress: tuple[Progress, ...]
    conflicts: tuple[CourseConflict, ...] = ()
    unclassified_courses: tuple[CompletedCourse, ...] = ()


class _TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.rows: list[list[str]] = []
        self._row: list[str] | None = None
        self._cell: list[str] | None = None

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag == "tr":
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
        elif tag == "tr" and self._row is not None:
            if self._row:
                self.rows.append(self._row)
            self._row = None


def _is_passing_grade(value: str) -> bool:
    normalized = value.strip()
    try:
        return float(normalized) >= 60
    except ValueError:
        return normalized in {"优秀", "良好", "中等", "及格", "合格", "通过", "免修"}


def _is_in_progress_grade(value: str) -> bool:
    """Recognise only explicit absence of a final grade, never a failing grade."""
    return value.strip() in {"", "-", "--", "—", "暂无", "未录入"}


def parse_grade_html(html: str) -> tuple[AcademicRecord, ...]:
    """Extract completion facts from one /cjcx/queryQmcj HTML page."""
    parser = _TableParser()
    parser.feed(html)
    header_index = next(
        (index for index, row in enumerate(parser.rows) if "课程代码" in row), None
    )
    if header_index is None:
        return ()
    headers = parser.rows[header_index]
    required = (
        "学年学期",
        "课程代码",
        "课程名称",
        "课程性质",
        "课程类别",
        "学分",
        "最终成绩",
    )
    if any(field not in headers for field in required):
        raise ValueError("成绩表缺少课程字段")
    positions = {field: headers.index(field) for field in required}
    records: list[AcademicRecord] = []
    for row in parser.rows[header_index + 1 :]:
        if len(row) <= max(positions.values()):
            continue
        try:
            records.append(
                AcademicRecord(
                    semester=row[positions["学年学期"]],
                    code=row[positions["课程代码"]],
                    name=row[positions["课程名称"]],
                    nature=row[positions["课程性质"]],
                    category=row[positions["课程类别"]],
                    credits=float(row[positions["学分"]]),
                    passed=_is_passing_grade(row[positions["最终成绩"]]),
                    in_progress=_is_in_progress_grade(row[positions["最终成绩"]]),
                )
            )
        except (TypeError, ValueError):
            continue
    return tuple(records)


def evaluate_progress(
    records: Iterable[AcademicRecord], baseline: RequirementBaseline
) -> ProgressReport:
    """Evaluate confirmed and selected-without-grade courses separately."""
    all_records = tuple(records)
    by_identity: dict[str, list[AcademicRecord]] = defaultdict(list)
    for record in all_records:
        if not record.passed:
            continue
        identity = record.code.strip() or " ".join(record.name.lower().split())
        by_identity[identity].append(record)

    unique: list[AcademicRecord] = []
    conflicts: list[CourseConflict] = []
    for identity, matches in by_identity.items():
        facts = {
            (
                " ".join(record.name.split()),
                record.nature.strip(),
                record.category.strip(),
                record.credits,
            )
            for record in matches
        }
        if len(facts) > 1:
            conflicts.append(CourseConflict(identity, tuple(matches)))
            continue
        unique.append(matches[0])

    grouped: dict[str, list[CompletedCourse]] = defaultdict(list)
    unclassified: list[CompletedCourse] = []
    for record in unique:
        if record.nature.strip() == "必修":
            continue
        course = CompletedCourse(
            record.code,
            record.name,
            record.nature,
            record.category,
            record.credits,
        )
        requirement_key = baseline.category_mapping.get(record.category.strip())
        if requirement_key is None:
            unclassified.append(course)
            continue
        grouped[requirement_key].append(course)

    # Current courses with an explicitly blank final-grade field are selected
    # facts. They contribute to estimates, but never to confirmed completion.
    in_progress_grouped: dict[str, list[CompletedCourse]] = defaultdict(list)
    in_progress_seen: set[str] = set()
    for record in all_records:
        if not record.in_progress or record.nature.strip() == "必修":
            continue
        identity = record.code.strip() or " ".join(record.name.lower().split())
        if identity in by_identity or identity in in_progress_seen:
            continue
        requirement_key = baseline.category_mapping.get(record.category.strip())
        if requirement_key is None:
            continue
        in_progress_seen.add(identity)
        in_progress_grouped[requirement_key].append(
            CompletedCourse(record.code, record.name, record.nature, record.category, record.credits)
        )

    progress_items: list[Progress] = []
    for requirement in baseline.requirements:
        contribution_keys = requirement.contribution_keys or (requirement.key,)
        matched = tuple(
            course
            for key in dict.fromkeys(contribution_keys)
            for course in grouped.get(key, ())
        )
        selected = tuple(
            course
            for key in dict.fromkeys(contribution_keys)
            for course in in_progress_grouped.get(key, ())
        )
        progress_items.append(
            Progress(
                requirement,
                sum(course.credits for course in matched),
                matched,
                selected,
            )
        )
    progress = tuple(progress_items)
    return ProgressReport(
        baseline.version, progress, tuple(conflicts), tuple(unclassified)
    )


_TABLE_REQUIREMENTS = {
    "本专业选修": "major_elective",
    "外专业选修": "outside_major_elective",
    "跨专业发展课程": "outside_major_elective",
    "文化素质课程": "cultural_quality",
}


def _credits(text: str) -> float:
    match = re.search(r"(\d+(?:\.\d+)?)\s*学分", text)
    if not match:
        raise ValueError(f"未找到学分要求: {text}")
    return float(match.group(1))


def parse_requirements(markdown: str | Path) -> tuple[Requirement, ...]:
    """Parse the small, stable requirement table and prose rules from the guide."""
    text = Path(markdown).read_text(encoding="utf-8") if isinstance(markdown, Path) else markdown
    found: dict[str, Requirement] = {}

    for line in text.splitlines():
        if not line.startswith("|"):
            continue
        cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
        if len(cells) < 2:
            continue
        key = _TABLE_REQUIREMENTS.get(cells[0])
        if key and re.search(r"\d+(?:\.\d+)?\s*学分", cells[1]):
            found[key] = Requirement(key, cells[0], _credits(cells[1]))

    cultural = re.search(r"文化素质课程.*?毕业前最低要求为\s*\*\*(\d+(?:\.\d+)?)\s*学分", text, re.DOTALL)
    if cultural:
        found["cultural_quality"] = Requirement("cultural_quality", "文化素质课程", float(cultural.group(1)))

    combined = re.search(r"两类学分合计要求\s*\*\*不少于\s*(\d+(?:\.\d+)?)\s*学分", text)
    innovation = re.search(r"创新创业学分\s*\*\*不少于\s*(\d+(?:\.\d+)?)\s*学分", text)
    practice = re.search(r"社会实践学分要求毕业前至少修满\s*\*\*(\d+(?:\.\d+)?)\s*学分", text)
    if combined:
        found["innovation_and_practice"] = Requirement(
            "innovation_and_practice",
            "创新创业 + 社会实践",
            float(combined.group(1)),
            ("innovation", "social_practice"),
        )
    if innovation:
        found["innovation"] = Requirement("innovation", "创新创业", float(innovation.group(1)))
    if practice:
        found["social_practice"] = Requirement("social_practice", "社会实践", float(practice.group(1)))

    return tuple(found.values())


def _bucket(course: CompletedCourse) -> str | None:
    category = course.category.strip()
    if category == "文理通识-文化素质教育课":
        return "cultural_quality"
    if category == "创新研修课":
        return "innovation"
    if category == "本专业选修":
        return "major_elective"
    if category in {"外专业选修", "跨专业发展课程"}:
        return "outside_major_elective"
    if category == "社会实践":
        return "social_practice"
    return None


def calculate_progress(
    requirements: Iterable[Requirement], courses: Iterable[CompletedCourse]
) -> tuple[Progress, ...]:
    """Exclude mandatory courses, classify the rest, and calculate deficits."""
    grouped: dict[str, list[CompletedCourse]] = defaultdict(list)
    for course in courses:
        if course.nature.strip() == "必修":
            continue
        bucket = _bucket(course)
        if bucket:
            grouped[bucket].append(course)

    result = []
    for requirement in requirements:
        contribution_keys = requirement.contribution_keys or (requirement.key,)
        matched = tuple(
            course
            for key in dict.fromkeys(contribution_keys)
            for course in grouped.get(key, ())
        )
        result.append(
            Progress(requirement, sum(course.credits for course in matched), matched)
        )
    return tuple(result)
