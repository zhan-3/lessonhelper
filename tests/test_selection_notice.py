import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from course_selection.notice import (
    confirm_notice,
    fetch_notice_text,
    load_notice,
    parse_notice,
    save_notice,
    update_notice,
)


NOTICE = """关于2026年秋季学期文化素质教育课程选课的通知
面向2025级本科生，选课时间为2026年8月26日08:00至2026年8月28日23:00。
请进入学生选课菜单完成操作。每人限选一门。"""


class SelectionNoticeTests(unittest.TestCase):
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
