import tempfile
import unittest
from pathlib import Path

from course_progress.collector import (
    AcademicRecord,
    CollectionCheckpoint,
    SemesterOption,
    SessionExpiredError,
    collect_grade_records,
    find_academic_frame,
    grade_query_parameters,
    load_checkpoint,
    resolve_academic_url,
    save_checkpoint,
)


def grade_page(code: str, semester: str, page_count: int) -> str:
    return f"""
    <table><tr><th>学年学期</th><th>课程代码</th><th>课程名称</th>
    <th>课程性质</th><th>课程类别</th><th>学分</th><th>最终成绩</th></tr>
    <tr><td>{semester}</td><td>{code}</td><td>测试课程</td><td>任选</td>
    <td>文理通识-文化素质教育课</td><td>1.0</td><td>合格</td></tr></table>
    <input type="hidden" id="pageCount" value="{page_count}">
    """


class CollectorTests(unittest.TestCase):
    def test_finds_academic_frame_in_any_open_tab(self):
        academic_frame = type("FakeFrame", (), {"url": "/cjcx/queryQmcj"})()

        class FakePage:
            def __init__(self, frame):
                self._frame = frame

            def frame(self, *, name):
                return self._frame if name == "iframename" else None

        self.assertIs(
            find_academic_frame(
                (FakePage(None), FakePage(academic_frame)),
            ),
            academic_frame,
        )

    def test_resolves_endpoint_inside_webvpn_application_prefix(self):
        current = (
            "https://webvpn.hitwh.edu.cn/http/"
            "77726476706e69737468656265737421fae0558f693861446900c7a99c406d3667/"
            "jdcx/queryXsjdcx"
        )

        result = resolve_academic_url(current, "/cjcx/queryQmcj")

        self.assertEqual(
            result,
            "https://webvpn.hitwh.edu.cn/http/"
            "77726476706e69737468656265737421fae0558f693861446900c7a99c406d3667/"
            "cjcx/queryQmcj",
        )

    def test_resolves_endpoint_inside_https_webvpn_application_prefix(self):
        current = (
            "https://webvpn.hitwh.edu.cn/https/"
            "77726476706e69737468656265737421f9e15192693861446900c7a99c406d36e9/"
            "portal/#!/service"
        )

        result = resolve_academic_url(current, "/cjcx/queryQmcj")

        self.assertEqual(
            result,
            "https://webvpn.hitwh.edu.cn/https/"
            "77726476706e69737468656265737421f9e15192693861446900c7a99c406d36e9/"
            "cjcx/queryQmcj",
        )

    def test_grade_query_parameters_match_initial_portal_query(self):
        self.assertEqual(
            grade_query_parameters("2025-20262", 1, page_size=20),
            {
                "pageXnxq": "2025-20262",
                "pageBkcxbj": "",
                "pageSfjg": "",
                "pageKcmc": "",
            },
        )
        self.assertEqual(
            grade_query_parameters("2025-20262", 2, page_size=20)["pageNo"],
            "2",
        )

    def test_collects_every_page_for_each_dynamic_semester(self):
        calls = []
        page_events = []

        def fetch_page(semester: str, page_number: int) -> str:
            calls.append((semester, page_number))
            return grade_page(
                f"{semester}-{page_number}", semester, 2 if page_number == 1 else 2
            )

        result = collect_grade_records(
            (
                SemesterOption("2025-20261", "2025秋季"),
                SemesterOption("2025-20262", "2026春季"),
            ),
            fetch_page,
            on_page=lambda semester, page_number, record_count: page_events.append(
                (semester.label, page_number, record_count)
            ),
        )

        self.assertEqual(
            calls,
            [
                ("2025-20261", 1),
                ("2025-20261", 2),
                ("2025-20262", 1),
                ("2025-20262", 2),
            ],
        )
        self.assertEqual(len(result.records), 4)
        self.assertEqual(result.failures, ())
        self.assertEqual(
            page_events,
            [
                ("2025秋季", 1, 1),
                ("2025秋季", 2, 1),
                ("2026春季", 1, 1),
                ("2026春季", 2, 1),
            ],
        )

    def test_failed_semester_is_reported_as_incomplete(self):
        def fetch_page(semester: str, page_number: int) -> str:
            raise RuntimeError("session expired")

        result = collect_grade_records(
            (SemesterOption("2025-20261", "2025秋季"),), fetch_page
        )

        self.assertEqual(result.records, ())
        self.assertFalse(result.complete)
        self.assertEqual(result.failures[0].semester_label, "2025秋季")
        self.assertEqual(result.failures[0].page_number, 1)

    def test_login_page_is_not_misread_as_empty_semester(self):
        result = collect_grade_records(
            (SemesterOption("2025-20261", "2025秋季"),),
            lambda semester, page_number: "<html><title>统一身份认证</title></html>",
        )

        self.assertFalse(result.complete)
        self.assertIn("统一身份认证", result.failures[0].message)

    def test_session_expiry_retries_page_after_reauthentication(self):
        calls = []
        reauthentications = []

        def fetch_page(semester: str, page_number: int) -> str:
            calls.append((semester, page_number))
            if len(calls) == 1:
                raise SessionExpiredError("登录已失效")
            return grade_page("MATH101", semester, 1)

        result = collect_grade_records(
            (SemesterOption("2025-20261", "2025秋季"),),
            fetch_page,
            on_session_expired=lambda: reauthentications.append(True),
        )

        self.assertEqual(calls, [("2025-20261", 1), ("2025-20261", 1)])
        self.assertEqual(reauthentications, [True])
        self.assertTrue(result.complete)
        self.assertEqual(len(result.records), 1)

    def test_resume_skips_completed_pages_from_checkpoint(self):
        semester = SemesterOption("2025-20261", "2025秋季")
        first_page_record = AcademicRecord(
            semester.value,
            "MATH101",
            "高等数学",
            "任选",
            "文理通识-文化素质教育课",
            1.0,
            True,
        )
        calls = []

        result = collect_grade_records(
            (semester,),
            lambda value, page: calls.append((value, page))
            or grade_page("MATH102", value, 2),
            checkpoint=CollectionCheckpoint(
                records=(first_page_record,),
                completed_pages=((semester.value, 1),),
                page_counts=((semester.value, 2),),
            ),
        )

        self.assertEqual(calls, [(semester.value, 2)])
        self.assertTrue(result.complete)
        self.assertEqual(
            {record.code for record in result.records}, {"MATH101", "MATH102"}
        )

    def test_checkpoint_round_trip_contains_no_scores(self):
        checkpoint = CollectionCheckpoint(
            records=(
                AcademicRecord(
                    "2025-20261",
                    "MATH101",
                    "高等数学",
                    "任选",
                    "文理通识-文化素质教育课",
                    1.0,
                    True,
                ),
            ),
            completed_pages=(("2025-20261", 1),),
            page_counts=(("2025-20261", 1),),
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "checkpoint.json"
            save_checkpoint(path, checkpoint)
            raw = path.read_text("utf-8")
            restored = load_checkpoint(path)

        self.assertEqual(restored, checkpoint)
        self.assertNotIn("成绩", raw)


if __name__ == "__main__":
    unittest.main()
