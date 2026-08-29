"""Local student facts used to match academic rules and selection notices."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, replace
from datetime import datetime, timezone
from pathlib import Path

DEFAULT_STUDENT_PROFILE_PATH = Path(
    ".private/academic-selection/student-profile.json"
)


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def normalize_grade(value: str) -> str:
    grade = value.strip().removesuffix("级")
    if not grade.isdigit() or len(grade) != 4:
        raise ValueError("年级必须是四位入学年份，例如 2025级")
    return grade


@dataclass(frozen=True)
class StudentProfile:
    grade: str
    major: str = ""
    academic_level: str = ""
    campus: str = ""
    updated_at: str = ""


def create_student_profile(
    *,
    grade: str,
    major: str = "",
    academic_level: str = "",
    campus: str = "",
) -> StudentProfile:
    return StudentProfile(
        grade=normalize_grade(grade),
        major=major.strip(),
        academic_level=academic_level.strip(),
        campus=campus.strip(),
        updated_at=_now(),
    )


def update_student_profile(profile: StudentProfile, **fields: str) -> StudentProfile:
    allowed = {"grade", "major", "academic_level", "campus"}
    unknown = set(fields) - allowed
    if unknown:
        raise ValueError(f"不支持的学生画像字段：{', '.join(sorted(unknown))}")
    values = {key: value.strip() for key, value in fields.items()}
    if "grade" in values:
        values["grade"] = normalize_grade(values["grade"])
    return replace(profile, **values, updated_at=_now())


def save_student_profile(path: Path, profile: StudentProfile) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(asdict(profile), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    temporary.replace(path)


def load_student_profile(path: Path) -> StudentProfile:
    data = json.loads(path.read_text(encoding="utf-8"))
    profile = StudentProfile(**data)
    normalize_grade(profile.grade)
    return profile
