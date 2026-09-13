"""Calculate course progress from raw passed-course records and guidance Markdown."""

from __future__ import annotations

import re
from collections import defaultdict
from collections.abc import Iterable, Mapping
from dataclasses import asdict, dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Literal, cast

from .baselines import (
    CONSTRAINT_KINDS,
    EVIDENCE_KINDS,
    ConstraintKind,
    EvidenceKind,
)

ConditionState = Literal["satisfied", "not_satisfied", "unknown"]
AssessmentReason = Literal[
    "confirmed_minimum_met",
    "confirmed_below_minimum",
    "grade_data_incomplete",
    "recognized_credit_coverage_missing",
    "classification_evidence_missing",
    "outside_major_track_unconfirmed",
    "required_subconstraint_unknown",
    "required_subconstraint_not_satisfied",
]

_REASON_DETAILS: dict[AssessmentReason, str] = {
    "confirmed_minimum_met": "学校成绩记录已确认达到最低值",
    "confirmed_below_minimum": "完整成绩记录中的已确认贡献低于最低值",
    "grade_data_incomplete": "成绩读取不完整，不能判断缺口",
    "recognized_credit_coverage_missing": "活动或认定学分尚未纳入数据覆盖",
    "classification_evidence_missing": "缺少明确的课程分类依据",
    "outside_major_track_unconfirmed": "尚未确认外专业课程所属的单一体系",
    "required_subconstraint_unknown": "必要子约束仍有未知项",
    "required_subconstraint_not_satisfied": "至少一项必要子约束未满足",
}


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


@dataclass(frozen=True)
class Requirement:
    key: str
    label: str
    minimum_credits: float
    contribution_keys: tuple[str, ...] = ()
    unit: Literal["credits", "courses"] = "credits"
    source: str = ""
    parent: str = ""
    constraint: ConstraintKind = ""
    evidence: EvidenceKind = "grade_records"
    required_conditions: tuple[str, ...] = ()


@dataclass(frozen=True)
class RequirementBaseline:
    version: str
    requirements: tuple[Requirement, ...]
    category_mapping: Mapping[str, str | tuple[str, ...]]


@dataclass(frozen=True)
class Progress:
    requirement: Requirement
    completed_credits: float
    courses: tuple[CompletedCourse, ...]

    @property
    def completed_amount(self) -> float:
        if self.requirement.unit == "courses":
            return float(len(self.courses))
        return self.completed_credits

    @property
    def remaining_credits(self) -> float:
        return max(0.0, self.requirement.minimum_credits - self.completed_amount)


@dataclass(frozen=True)
class RequirementAssessment:
    progress: Progress
    state: ConditionState
    reason: AssessmentReason

    @property
    def confirmed_amount(self) -> float:
        return self.progress.completed_amount

    @property
    def confirmed_gap(self) -> float:
        return self.progress.remaining_credits


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
                )
            )
        except (TypeError, ValueError):
            continue
    return tuple(records)


def baseline_from_definition(definition: Mapping[str, object] | None) -> RequirementBaseline:
    """Create the calculator baseline from one immutable serialized definition."""
    if not isinstance(definition, Mapping):
        raise TypeError("requirement baseline definition is missing")
    raw_requirements = definition.get("requirements")
    raw_mapping = definition.get("category_mapping")
    if not isinstance(raw_requirements, (list, tuple)) or not isinstance(raw_mapping, Mapping):
        raise TypeError("requirement baseline definition is incomplete")
    requirements = []
    for item in raw_requirements:
        if not isinstance(item, Mapping):
            raise TypeError("requirement definition must be an object")
        unit = str(item.get("unit", "credits"))
        constraint = str(item.get("constraint", ""))
        evidence = str(item.get("evidence", "grade_records"))
        if unit not in {"credits", "courses"}:
            raise ValueError(f"unknown requirement unit: {unit}")
        if constraint not in CONSTRAINT_KINDS:
            raise ValueError(f"unknown requirement constraint: {constraint}")
        if evidence not in EVIDENCE_KINDS:
            raise ValueError(f"unknown requirement evidence: {evidence}")
        requirements.append(Requirement(
            key=str(item.get("key", "")),
            label=str(item.get("label", "")),
            minimum_credits=float(item.get("minimum", 0)),
            contribution_keys=tuple(str(value) for value in item.get("contribution_keys", ())),
            unit=cast(Literal["credits", "courses"], unit),
            source=str(item.get("source", "")),
            parent=str(item.get("parent", "")),
            constraint=cast(ConstraintKind, constraint),
            evidence=cast(EvidenceKind, evidence),
            required_conditions=tuple(str(value) for value in item.get("required_conditions", ())),
        ))
    category_mapping: dict[str, str | tuple[str, ...]] = {}
    for key, value in raw_mapping.items():
        if isinstance(value, str):
            category_mapping[str(key)] = value
        elif isinstance(value, (list, tuple)) and value:
            category_mapping[str(key)] = tuple(str(target) for target in value)
        else:
            raise TypeError("category mapping target must be a string or non-empty list")
    return RequirementBaseline(
        version=str(definition.get("version", "")),
        requirements=tuple(requirements),
        category_mapping=category_mapping,
    )


def evaluate_progress(
    records: Iterable[AcademicRecord], baseline: RequirementBaseline
) -> ProgressReport:
    """Evaluate confirmed course progress against one requirement baseline."""
    by_identity: dict[str, list[AcademicRecord]] = defaultdict(list)
    for record in records:
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
        mapping_value = baseline.category_mapping.get(record.category.strip())
        if mapping_value is None:
            unclassified.append(course)
            continue
        targets = (mapping_value,) if isinstance(mapping_value, str) else mapping_value
        for requirement_key in dict.fromkeys(targets):
            grouped[requirement_key].append(course)

    progress_items: list[Progress] = []
    for requirement in baseline.requirements:
        contribution_keys = (
            requirement.key,
            *requirement.contribution_keys,
            *(item.key for item in baseline.requirements if item.parent == requirement.key),
        )
        matched_by_identity = {
            course.code.strip() or " ".join(course.name.lower().split()): course
            for key in dict.fromkeys(contribution_keys)
            for course in grouped.get(key, ())
        }
        matched = tuple(matched_by_identity.values())
        progress_items.append(
            Progress(
                requirement,
                sum(course.credits for course in matched),
                matched,
            )
        )
    progress = tuple(progress_items)
    return ProgressReport(
        baseline.version, progress, tuple(conflicts), tuple(unclassified)
    )


def assess_progress(
    report: ProgressReport, *, data_complete: bool
) -> tuple[RequirementAssessment, ...]:
    """Classify confirmed minimums without treating uncovered facts as zero."""
    raw: dict[str, RequirementAssessment] = {}
    for progress in report.progress:
        requirement = progress.requirement
        amount = progress.completed_amount
        if amount >= requirement.minimum_credits:
            if requirement.constraint == "single_track":
                state, reason = "unknown", "outside_major_track_unconfirmed"
            else:
                state, reason = "satisfied", "confirmed_minimum_met"
        elif not data_complete:
            state, reason = "unknown", "grade_data_incomplete"
        elif requirement.evidence == "grade_and_recognition":
            state, reason = "unknown", "recognized_credit_coverage_missing"
        elif requirement.evidence == "classification":
            state, reason = "unknown", "classification_evidence_missing"
        else:
            state, reason = "not_satisfied", "confirmed_below_minimum"
        raw[requirement.key] = RequirementAssessment(progress, state, reason)

    assessed = dict(raw)
    for key, item in raw.items():
        conditions = item.progress.requirement.required_conditions
        children = [raw[condition] for condition in conditions if condition in raw]
        if any(child.state == "not_satisfied" for child in children):
            assessed[key] = RequirementAssessment(
                item.progress, "not_satisfied", "required_subconstraint_not_satisfied"
            )
        elif item.state == "satisfied" and any(child.state == "unknown" for child in children):
            assessed[key] = RequirementAssessment(
                item.progress, "unknown", "required_subconstraint_unknown"
            )
    return tuple(assessed[item.requirement.key] for item in report.progress)


def apply_recognized_credit_estimates(
    progress_items: list[dict[str, object]],
    baseline_definition: Mapping[str, object],
    declarations: Iterable[Mapping[str, object]],
) -> list[dict[str, object]]:
    """Add local declaration estimates without changing confirmed school facts."""
    confirmed_identities = {
        str(course.get("code") or " ".join(str(course.get("name", "")).lower().split()))
        for item in progress_items
        for course in item.get("courses", ())
        if isinstance(course, Mapping)
    }
    unique_declarations = {
        str(item.get("identity", "")): item for item in declarations if item.get("identity")
    }
    claimed_links: set[str] = set()
    declaration_overlap: dict[str, tuple[bool, str]] = {}
    for identity, declaration in unique_declarations.items():
        linked_identity = str(declaration.get("linked_course_identity", ""))
        if linked_identity and linked_identity in confirmed_identities:
            declaration_overlap[identity] = (False, "linked_confirmed_course")
        elif linked_identity and linked_identity in claimed_links:
            declaration_overlap[identity] = (False, "duplicate_declaration_link")
        else:
            declaration_overlap[identity] = (
                True,
                "linked_identity_not_confirmed" if linked_identity else "unverified_overlap",
            )
            if linked_identity:
                claimed_links.add(linked_identity)
    requirements = baseline_definition.get("requirements", ())
    contribution_targets = {
        str(item.get("key", "")): {
            str(item.get("key", "")),
            *(str(key) for key in item.get("contribution_keys", ())),
        }
        for item in requirements
        if isinstance(item, Mapping)
    }
    result: list[dict[str, object]] = []
    for progress in progress_items:
        key = str(progress.get("key", ""))
        applied: list[dict[str, object]] = []
        declared_amount = 0.0
        for declaration in unique_declarations.values():
            if str(declaration.get("category", "")) not in contribution_targets.get(key, {key}):
                continue
            contributes, overlap_status = declaration_overlap[
                str(declaration.get("identity", ""))
            ]
            credits = float(declaration.get("credits", 0))
            if contributes:
                declared_amount += credits
            applied.append({
                **dict(declaration),
                "contributes": contributes,
                "overlap_status": overlap_status,
            })
        confirmed_amount = float(progress.get("confirmed_amount", 0))
        minimum = float(progress.get("minimum", 0))
        estimated_amount = confirmed_amount + declared_amount
        result.append({
            **progress,
            "declared_amount": round(declared_amount, 6),
            "estimated_amount": round(estimated_amount, 6),
            "estimated_gap": round(max(0.0, minimum - estimated_amount), 6),
            "declarations": applied,
            "manual_review_required": any(
                item["overlap_status"] not in {
                    "linked_confirmed_course", "duplicate_declaration_link",
                }
                for item in applied
            ),
        })
    return result


def apply_course_label_estimates(
    progress_items: list[dict[str, object]],
    unclassified_courses: Iterable[Mapping[str, object]],
    labels: Iterable[Mapping[str, object]],
    selected_track: str,
) -> list[dict[str, object]]:
    """Apply user course labels to estimates without rewriting confirmed facts."""
    course_by_identity: dict[str, Mapping[str, object]] = {}
    for item in progress_items:
        for course in item.get("courses", ()):
            if isinstance(course, Mapping):
                identity = str(
                    course.get("code")
                    or " ".join(str(course.get("name", "")).lower().split())
                )
                course_by_identity[identity] = course
    for course in unclassified_courses:
        identity = str(
            course.get("code")
            or " ".join(str(course.get("name", "")).lower().split())
        )
        course_by_identity[identity] = course
    labels_by_identity = {
        str(label.get("course_identity", "")): label
        for label in labels
        if str(label.get("course_identity", "")) in course_by_identity
    }
    result: list[dict[str, object]] = []
    for item in progress_items:
        key = str(item.get("key", ""))
        estimated_amount = float(item.get("estimated_amount", item.get("confirmed_amount", 0)))
        labeled: list[dict[str, object]] = []
        estimate_replaces_confirmed = False
        if key in {"cultural_quality_d", "four_histories"}:
            field = "d_category" if key == "cultural_quality_d" else "four_histories"
            confirmed_ids = {
                str(course.get("code") or " ".join(str(course.get("name", "")).lower().split()))
                for course in item.get("courses", ())
                if isinstance(course, Mapping)
            }
            matching = {
                identity: course_by_identity[identity]
                for identity, label in labels_by_identity.items()
                if label.get(field) is True and identity not in confirmed_ids
            }
            labeled = [
                {"course_identity": identity, **dict(course), "source": "user_declared"}
                for identity, course in matching.items()
            ]
            addition = (
                float(len(matching)) if key == "four_histories"
                else sum(float(course.get("credits", 0)) for course in matching.values())
            )
            estimated_amount += addition
        elif key == "outside_major_elective":
            estimate_replaces_confirmed = True
            matching = {
                identity: course_by_identity[identity]
                for identity, label in labels_by_identity.items()
                if selected_track and label.get("outside_track") == selected_track
            }
            labeled = [
                {"course_identity": identity, **dict(course), "source": "user_declared"}
                for identity, course in matching.items()
            ]
            estimated_amount = sum(
                float(course.get("credits", 0)) for course in matching.values()
            )
        minimum = float(item.get("minimum", 0))
        estimated_gap = max(0.0, minimum - estimated_amount)
        updated = {
            **item,
            "estimated_amount": round(estimated_amount, 6),
            "estimated_gap": round(estimated_gap, 6),
            "labeled_courses": labeled,
            "estimate_replaces_confirmed": estimate_replaces_confirmed,
        }
        if key in {"cultural_quality_d", "four_histories", "outside_major_elective"}:
            updated["estimated_condition_status"] = (
                "estimated_satisfied" if estimated_gap == 0 else "unknown"
            )
        if key == "outside_major_elective":
            updated["selected_track"] = selected_track
            updated["other_track_course_identities"] = [
                identity for identity, label in labels_by_identity.items()
                if label.get("outside_track") and label.get("outside_track") != selected_track
            ]
            outside_identities = {
                str(course.get("code") or " ".join(str(course.get("name", "")).lower().split()))
                for course in item.get("courses", ())
                if isinstance(course, Mapping)
            }
            updated["unknown_track_course_identities"] = sorted(
                identity for identity in outside_identities
                if not labels_by_identity.get(identity, {}).get("outside_track")
            )
        result.append(updated)
    return result


def confirmed_progress_items(
    report: ProgressReport, *, data_complete: bool
) -> list[dict[str, object]]:
    """Serialize every baseline requirement using the confirmed-evidence vocabulary."""
    def rule_detail(requirement: Requirement) -> str:
        if requirement.source == "manual-supplement":
            return "人工补充参考规则：不代表所有年级的统一正式要求，请核对个人培养方案。"
        if requirement.constraint == "single_track":
            return "总额之外还须确认课程来自同一个已选体系。"
        if requirement.parent:
            return "这是总额内的必要子约束，不额外重复计入总量。"
        return ""

    return [
        {
            "key": item.progress.requirement.key,
            "label": item.progress.requirement.label,
            "unit": item.progress.requirement.unit,
            "source": item.progress.requirement.source,
            "parent": item.progress.requirement.parent,
            "constraint": item.progress.requirement.constraint,
            "rule_detail": rule_detail(item.progress.requirement),
            "minimum": item.progress.requirement.minimum_credits,
            "confirmed_amount": item.confirmed_amount,
            "confirmed_gap": item.confirmed_gap,
            "condition_status": item.state,
            "condition_detail": {
                "satisfied": "已满足", "not_satisfied": "未满足", "unknown": "未知",
            }[item.state],
            "reason": item.reason,
            "reason_detail": _REASON_DETAILS[item.reason],
            # Retain the established fields while clients migrate to the
            # evidence-aware vocabulary above.
            "required_credits": item.progress.requirement.minimum_credits,
            "completed_credits": item.progress.completed_credits,
            "remaining_credits": item.confirmed_gap,
            "courses": [asdict(course) for course in item.progress.courses],
        }
        for item in assess_progress(report, data_complete=data_complete)
    ]


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
