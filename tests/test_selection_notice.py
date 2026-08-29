import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from course_selection.notice import (
    confirm_notice,
    fetch_notice_text,
    load_notice,
    notice_selection_categories,
    notice_semester_label,
    parse_notice,
    parse_selection_windows,
    save_notice,
    update_notice,
)

NOTICE = """关于2026年秋季学期文化素质教育课程选课的通知
面向2025级本科生，选课时间为2026年8月26日08:00至2026年8月28日23:00。
请进入学生选课菜单完成操作。每人限选一门。"""

SCHEDULE_NOTICE = """首页
关于2026年秋季学期各类课程选课时间安排的通知
2026年9月4日8:30-2026年9月4日17:00
2026级
大学体育
新教务系统
2026年8月29日10:30-2026年8月30日17:00
2024级、2025级
大学体育
新教务系统
2026年8月29日8:30-2026年8月29日11:00
2023级、2024级、
2025级
创新研修课、创新实验课、创新创业课——选课
新教务系统
2026年8月29日14:30-2026年8月29日17:00
2023级、2024级、2025级
创新研修课、创新实验课、创新创业课——退课
新教务系统
2026年8月29日14:30-2026年8月29日16:30
2025级
文化素质教育课——选课
新教务系统
2026年8月29日8:30-2026年8月30日17:00
2023级、2024级、2025级
跨专业发展课程
新教务系统
2026年8月29日8:30-2026年8月30日17:00
2023级、2024级、2025级
补修
填写纸质申请表
报送院系教务员
"""


class SelectionNoticeTests(unittest.TestCase):
    def test_parses_multirow_schedule_and_filters_by_grade_action_and_method(self):
        notice = parse_notice(SCHEDULE_NOTICE)

        self.assertEqual(
            notice.title,
            "关于2026年秋季学期各类课程选课时间安排的通知",
        )
        self.assertEqual(len(parse_selection_windows(SCHEDULE_NOTICE)), 7)
        self.assertEqual(notice.selection_type, "多类别")
        self.assertEqual(notice_selection_categories(notice, grade="2026"), ("ty",))
        self.assertEqual(
            notice_selection_categories(notice, grade="2025"),
            ("ty", "cxyx", "cxsy", "cxcy", "szhx", "xsxk"),
        )

    def test_parses_selection_window_without_marking_it_confirmed(self):
        notice = parse_notice(
            NOTICE,
            source_url="https://jwc.example/notice/1",
        )

        self.assertEqual(notice.term, "2026年秋季学期")
        self.assertEqual(notice.selection_type, "文化素质")
        self.assertEqual(notice.opens_at, "2026年8月26日08:00")
        self.assertEqual(notice.closes_at, "2026年8月28日23:00")
        self.assertEqual(notice.status, "pending_confirmation")
        self.assertEqual(notice.missing_fields, ())
        self.assertEqual(notice_selection_categories(notice), ("szhx",))
        self.assertEqual(notice_semester_label(notice), "2026秋季")

    def test_notice_categories_include_only_explicitly_named_types(self):
        notice = parse_notice(
            "关于2026年秋季学期大学外语、体育课程选课的通知\n"
            "本轮不开放其他类别。"
        )

        self.assertEqual(notice_selection_categories(notice), ("yy", "ty"))

    def test_missing_required_fields_cannot_be_confirmed(self):
        notice = parse_notice("选课通知\n请关注后续安排", source_kind="manual")

        self.assertIn("term", notice.missing_fields)
        with self.assertRaisesRegex(ValueError, "选课通知仍缺少"):
            confirm_notice(notice)

    def test_user_can_fill_missing_fields_then_confirm_and_reload(self):
        notice = parse_notice("选课通知", source_kind="manual")
        completed = update_notice(
            notice,
            term="2026年秋季学期",
            selection_type="大学外语",
            opens_at="2026年8月26日08:00",
            closes_at="2026年8月26日22:00",
        )
        confirmed = confirm_notice(completed)

        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "selection-notice.json"
            save_notice(path, confirmed)
            self.assertEqual(load_notice(path), confirmed)

    def test_fetches_notice_text_from_an_http_link(self):
        class Headers:
            def get_content_charset(self):
                return "utf-8"

        class Response:
            headers = Headers()

            def read(self):
                return "<html><title>选课通知</title><p>2026年秋季学期</p></html>".encode()

            def __enter__(self):
                return self

            def __exit__(self, *_):
                return False

        with patch("course_selection.notice.urlopen", return_value=Response()):
            notice = parse_notice(fetch_notice_text("https://jwc.example/notice/1"))

        self.assertEqual(notice.title, "选课通知")
        self.assertEqual(notice.term, "2026年秋季学期")
