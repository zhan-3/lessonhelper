import io
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from course_selection.web import create_app


class SelectionWebTests(unittest.TestCase):
    def test_user_can_import_and_confirm_notice(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            client = create_app(Path(temp_dir)).test_client()
            response = client.post(
                "/notices",
                data={
                    "source_url": "https://jwc.example/notice/1",
                    "text": "2026年秋季学期文化素质教育课程选课通知\n选课时间：2026年8月26日08:00至2026年8月28日23:00",
                },
            )
            self.assertEqual(response.status_code, 302)
            response = client.post("/notices/confirm")
            self.assertEqual(response.status_code, 302)
            self.assertIn("confirmed", (Path(temp_dir) / "selection-notice.json").read_text("utf-8"))

    def test_dashboard_shows_missing_notice_fields_before_confirmation(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            client = create_app(Path(temp_dir)).test_client()
            client.post("/notices", data={"text": "选课通知"})

            dashboard = client.get("/")

            self.assertIn("待补充", dashboard.get_data(as_text=True))

    def test_user_can_upload_timetable_and_see_it_on_dashboard(self):
        workbook = Workbook()
        workbook.active.append(["学期", "课程名称", "星期", "节次", "周次"])
        workbook.active.append(["2026年秋季学期", "高等数学", "星期一", "1-2", "1-16"])
        stream = io.BytesIO()
        workbook.save(stream)
        stream.seek(0)

        with tempfile.TemporaryDirectory() as temp_dir:
            client = create_app(Path(temp_dir)).test_client()
            response = client.post(
                "/timetable",
                data={"timetable": (stream, "current.xlsx")},
                content_type="multipart/form-data",
            )
            self.assertEqual(response.status_code, 302)
            dashboard = client.get("/")
            self.assertEqual(dashboard.status_code, 200)
            self.assertIn("高等数学", dashboard.get_data(as_text=True))
            self.assertIn("current.xlsx", dashboard.get_data(as_text=True))

    def test_replacing_timetable_requires_explicit_confirmation(self):
        workbook = Workbook()
        workbook.active.append(["学期", "课程名称", "星期", "节次", "周次"])
        workbook.active.append(["2026年秋季学期", "高等数学", "星期一", "1-2", "1-16"])
        stream = io.BytesIO()
        workbook.save(stream)
        first_payload = stream.getvalue()

        with tempfile.TemporaryDirectory() as temp_dir:
            client = create_app(Path(temp_dir)).test_client()
            first = client.post(
                "/timetable",
                data={"timetable": (io.BytesIO(first_payload), "current.xlsx")},
                content_type="multipart/form-data",
            )
            self.assertEqual(first.status_code, 302)

            rejected = client.post(
                "/timetable",
                data={"timetable": (io.BytesIO(first_payload), "current.xlsx")},
                content_type="multipart/form-data",
            )
            self.assertEqual(rejected.status_code, 409)

    def test_dashboard_presents_read_only_selection_result(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            (root / "selection-entry.json").write_text(
                '{"status":"ready","request_url":"/selection/list",'
                '"sections":[{"name":"人工智能导论","credits":"2",'
                '"teacher":"李老师","time":"周一 1-2","capacity":"12",'
                '"selected":false}]}',
                encoding="utf-8",
            )
            dashboard = create_app(root).test_client().get("/")

            self.assertIn("人工智能导论", dashboard.get_data(as_text=True))
            self.assertIn("ready", dashboard.get_data(as_text=True))
