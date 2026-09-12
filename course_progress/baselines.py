"""Immutable requirement-baseline definitions exposed by the local workbench."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Any, Literal


@dataclass(frozen=True)
class BaselineRequirement:
    key: str
    label: str
    minimum: float
    unit: Literal["credits", "courses"] = "credits"
    source: str = ""
    parent: str = ""
    constraint: str = ""
    contribution_keys: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        value = asdict(self)
        return {key: item for key, item in value.items() if item not in {"", ()}}


@dataclass(frozen=True)
class RequirementBaselineDefinition:
    version: str
    title: str
    authority: str
    source_summary: str
    applicability: str
    coverage: str
    disclaimer: str
    requirements: tuple[BaselineRequirement, ...]
    category_mapping: tuple[tuple[str, str], ...]
    manual_supplements: tuple[str, ...] = ()
    requires_explicit_selection: bool = True

    def __post_init__(self) -> None:
        keys = {item.key for item in self.requirements}
        contribution_keys = {
            key for item in self.requirements for key in item.contribution_keys
        }
        if len(keys) != len(self.requirements):
            raise ValueError(f"duplicate requirement key in {self.version}")
        if any(item.minimum <= 0 for item in self.requirements):
            raise ValueError(f"non-positive requirement in {self.version}")
        if any(item.parent and item.parent not in keys for item in self.requirements):
            raise ValueError(f"unknown parent requirement in {self.version}")
        allowed_mapping_targets = keys | contribution_keys
        if any(target not in allowed_mapping_targets for _, target in self.category_mapping):
            raise ValueError(f"unknown category mapping target in {self.version}")

    def to_dict(self) -> dict[str, Any]:
        return {
            "version": self.version,
            "title": self.title,
            "authority": self.authority,
            "requires_explicit_selection": self.requires_explicit_selection,
            "source_summary": self.source_summary,
            "applicability": self.applicability,
            "coverage": self.coverage,
            "disclaimer": self.disclaimer,
            "manual_supplements": list(self.manual_supplements),
            "requirements": [item.to_dict() for item in self.requirements],
            "category_mapping": dict(self.category_mapping),
        }


_CATEGORY_MAPPING = (
    ("本专业选修", "major_elective"),
    ("外专业选修", "outside_major_elective"),
    ("跨专业发展课程", "outside_major_elective"),
    ("文理通识-文化素质教育课", "cultural_quality"),
    ("创新研修课", "innovation"),
    ("创新实验课", "innovation"),
    ("创新创业课程", "innovation"),
    ("创业课程", "innovation"),
    ("社会实践", "social_practice"),
)

_BASELINES = (
    RequirementBaselineDefinition(
        version="guide-2026",
        title="校园培养方案解读（2026 年版）",
        authority="extracted-guide",
        source_summary="项目内提纯指南；最终要求以个人培养方案为准。",
        applicability="参考指南没有证明其规则统一适用于所有年级和专业，必须由用户主动选择。",
        coverage="现有选修与认定学分总额，不包含完整专业培养方案。",
        disclaimer="规划参考，不是学校正式毕业审核结论。",
        requirements=(
            BaselineRequirement("major_elective", "本专业选修", 3.0),
            BaselineRequirement("outside_major_elective", "跨专业发展课程", 10.0),
            BaselineRequirement("cultural_quality", "文化素质课程", 8.0),
            BaselineRequirement(
                "innovation_and_practice", "创新创业 + 社会实践", 6.0,
                contribution_keys=("innovation", "social_practice"),
            ),
            BaselineRequirement("social_practice", "社会实践", 1.0),
        ),
        category_mapping=_CATEGORY_MAPPING,
    ),
    RequirementBaselineDefinition(
        version="basic-graduation-reference-v1",
        title="基础毕业要求参考基线 v1",
        authority="reference",
        source_summary="项目内提纯指南，加上用户确认的非官方参考图人工补充。",
        applicability=(
            "用户主动选择的个人参考口径，不自动适用于任何年级或专业。"
            "跨专业 10 学分在现有指南中仅明确覆盖 2022—2024 级；其他年级须核对个人培养方案。"
        ),
        coverage="基础选修及认定学分、必要合计与子约束；不含专业必修和完整培养方案审核。",
        disclaimer="个人规划参考，非学校正式培养方案，也非学校正式毕业审核结论。",
        manual_supplements=(
            "创新创业单项至少 4 学分尚未证实为所有年级统一正式规定，使用前需按个人培养方案核对。",
        ),
        requirements=(
            BaselineRequirement("major_elective", "本专业选修", 3.0, source="extracted-guide"),
            BaselineRequirement("innovation", "创新创业", 4.0, source="manual-supplement"),
            BaselineRequirement("social_practice", "社会实践", 1.0, source="extracted-guide"),
            BaselineRequirement(
                "innovation_and_practice", "创新创业 + 社会实践", 6.0,
                source="extracted-guide", contribution_keys=("innovation", "social_practice"),
            ),
            BaselineRequirement("cultural_quality", "文化素质课程", 8.0, source="extracted-guide"),
            BaselineRequirement(
                "cultural_quality_d", "文化素质 D 类", 2.0,
                source="extracted-guide", parent="cultural_quality",
            ),
            BaselineRequirement(
                "four_histories", "四史课程", 1, unit="courses",
                source="extracted-guide", parent="cultural_quality",
            ),
            BaselineRequirement(
                "outside_major_elective", "跨专业发展课程", 10.0,
                source="reference-scope", constraint="single_track",
            ),
        ),
        category_mapping=_CATEGORY_MAPPING,
    ),
)
_BASELINES_BY_VERSION = {item.version: item for item in _BASELINES}


def requirement_baselines() -> list[dict[str, Any]]:
    """Return serialized copies of every immutable baseline in display order."""
    return [item.to_dict() for item in _BASELINES]


def requirement_baseline(version: str) -> dict[str, Any] | None:
    """Return one immutable baseline definition as a detached API value."""
    baseline = _BASELINES_BY_VERSION.get(version)
    return baseline.to_dict() if baseline is not None else None
