"""Canonical academic course-category definitions shared by notice and discovery."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class CourseCategoryDefinition:
    code: str
    label: str
    notice_aliases: tuple[str, ...]
    menu_aliases: tuple[str, ...]
    navigation_priority: int


COURSE_CATEGORIES = (
    CourseCategoryDefinition(
        code="szhx",
        label="文化素质核心",
        notice_aliases=("文化素质核心", "文化素质", "素质教育"),
        menu_aliases=("文化素质核心", "素质教育"),
        navigation_priority=250,
    ),
    CourseCategoryDefinition(
        code="yy",
        label="英语",
        notice_aliases=("大学外语", "英语"),
        menu_aliases=("英语",),
        navigation_priority=205,
    ),
    CourseCategoryDefinition(
        code="ty",
        label="体育",
        notice_aliases=("体育",),
        menu_aliases=("体育", "体育选课"),
        navigation_priority=210,
    ),
    CourseCategoryDefinition(
        code="cxyx",
        label="创新研修",
        notice_aliases=("创新研修",),
        menu_aliases=("创新研修",),
        navigation_priority=232,
    ),
    CourseCategoryDefinition(
        code="cxsy",
        label="创新实验",
        notice_aliases=("创新实验",),
        menu_aliases=("创新实验课", "创新实验"),
        navigation_priority=230,
    ),
    CourseCategoryDefinition(
        code="cxcy",
        label="创新创业",
        notice_aliases=("创新创业",),
        menu_aliases=("创新创业",),
        navigation_priority=228,
    ),
    CourseCategoryDefinition(
        code="xsyt",
        label="新生研讨",
        notice_aliases=("新生研讨",),
        menu_aliases=("新生研讨",),
        navigation_priority=226,
    ),
    CourseCategoryDefinition(
        code="tsk",
        label="未来技术学院课程",
        notice_aliases=("未来技术学院",),
        menu_aliases=("未来技术学院课程",),
        navigation_priority=224,
    ),
    CourseCategoryDefinition(
        code="xsxk",
        label="外专业课程",
        notice_aliases=("外专业课程", "跨专业"),
        menu_aliases=("外专业课程", "跨专业选课"),
        navigation_priority=222,
    ),
    CourseCategoryDefinition(
        code="wzy",
        label="微专业选课",
        notice_aliases=("微专业",),
        menu_aliases=("微专业选课",),
        navigation_priority=120,
    ),
)

CATEGORY_BY_CODE = {item.code: item for item in COURSE_CATEGORIES}
CATEGORY_LABELS = {item.code: item.label for item in COURSE_CATEGORIES}
NOTICE_CATEGORY_PATTERNS = tuple(
    (alias, item.code)
    for item in COURSE_CATEGORIES
    for alias in item.notice_aliases
)
CATEGORY_MENU_KEYWORDS = tuple(
    alias
    for item in COURSE_CATEGORIES
    for alias in item.menu_aliases
)
