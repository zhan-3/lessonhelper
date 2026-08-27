import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from course_selection.timetable import import_timetable, timetable_snapshot_payload


class TimetableImportTests(unittest.TestCase):
    def _write_workbook(self, rows):
        workbook = Workbook()
        sheet = workbook.active
        for row in rows:
            sheet.append(row)
        temp_dir = tempfile.TemporaryDirectory()
        path = Path(temp_dir.name) / "timetable.xlsx"
        workbook.save(path)
        return temp_dir, path

    def test_imports_week_and_period_ranges_from_school_table(self):
        temp_dir, path = self._write_workbook(
            [
                ["学年学期", "课程代码", "课程名称", "星期", "节次", "周次", "单双周", "教室"],
                ["2026年秋季学期", "MATH101", "高等数学", "星期一", "1-2", "1-16周", "", "A101"],
            ]
        )
        self.addCleanup(temp_dir.cleanup)

        entries = import_timetable(path, expected_term="2026年秋季学期")

        self.assertEqual(len(entries), 1)
        self.assertEqual(entries[0].weekday, 1)
        self.assertEqual(entries[0].start_period, 1)
        self.assertEqual(entries[0].end_period, 2)
        self.assertEqual(entries[0].week_end, 16)
        self.assertEqual(entries[0].week_parity, "all")

    def test_rejects_missing_required_columns_and_term_mismatch(self):
        temp_dir, path = self._write_workbook(
            [["学期", "课程名称"], ["2026年秋季学期", "课程"]]
        )
        self.addCleanup(temp_dir.cleanup)

        with self.assertRaisesRegex(ValueError, "缺少必要列"):
            import_timetable(path)

    def test_rejects_term_mismatch_before_enabling_snapshot(self):
        temp_dir, path = self._write_workbook(
            [
                ["学期", "课程名称", "星期", "节次", "周次"],
                ["2026年秋季学期", "课程", "一", "1", "1-16"],
            ]
        )
        self.addCleanup(temp_dir.cleanup)

        with self.assertRaisesRegex(ValueError, "学期不匹配"):
            import_timetable(path, expected_term="2026年春季学期")

    def test_imports_school_grid_timetable_and_preserves_non_contiguous_weeks(self):
        temp_dir, path = self._write_workbook(
            [
                ["2026秋季学期(2025211052)张浩翔课表"],
                ["", "", "星期一", "星期二"],
                ["上午", "1-2", "马克思主义基本原理◇赵聪妹[1-5，7-16周]◇N楼-331", ""],
                ["下午", "5-6", "", "离散数学◇教师[10-17周]◇M楼-201"],
            ]
        )
        self.addCleanup(temp_dir.cleanup)

        entries = import_timetable(path, expected_term="2026年秋季学期")

        self.assertEqual(len(entries), 2)
        self.assertEqual(entries[0].course_name, "马克思主义基本原理")
        self.assertEqual(entries[0].week_numbers, (1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16))
        self.assertEqual(entries[0].location, "N楼-331")
        self.assertEqual(entries[1].weekday, 2)

    def test_pending_time_is_conflict_unknown_not_free(self):
        temp_dir, path = self._write_workbook(
            [
                ["学期", "课程名称", "星期", "节次", "周次"],
                ["2026年秋季学期", "待定课程", "星期一", "待定", "待定"],
            ]
        )
        self.addCleanup(temp_dir.cleanup)
        entries = import_timetable(path, expected_term="2026年秋季学期")
        self.assertEqual(entries[0].conflict_status, "unknown")
        self.assertIsNone(entries[0].start_period)
        self.assertIsNone(entries[0].week_start)
        payload = timetable_snapshot_payload(entries, source_name="current.xlsx")
        self.assertEqual(payload["source_kind"], "user-imported")
        self.assertEqual(payload["entries"][0]["conflict_status"], "unknown")
