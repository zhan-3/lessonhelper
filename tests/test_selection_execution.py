import json
import tempfile
import unittest
from pathlib import Path

from course_selection.persistence import WorkspaceDatabase
from course_selection.selection_execution import (
    VerifiedSelectionExecutionAdapter,
    classify_selection_execution_html,
)
from course_selection.workbench import create_workbench_app
from course_selection.workbench_service import WorkbenchService
from tests.test_workbench import FakeGateway

COURSE_HTML = """
<form id="queryform"><input name="token" value="fresh"><input name="pageXklb" value="xsxk"><input name="pageXnxq" value="2026-20271"></form>
<table><tr><th></th><th>课程代码</th><th>课程名称</th><th>上课信息</th></tr>
<tr><td><a onclick="saveXsxk1('TERM-TASK-001')">选课</a></td><td>C1</td><td>课程</td><td>星期一第1,2节</td></tr></table>
"""


class ResultClassificationTests(unittest.TestCase):
    def test_classifies_success_capacity_rejection_and_unknown(self):
        self.assertEqual("selected", classify_selection_execution_html("<script>alert('选课成功')</script>")[0])
        self.assertEqual("capacity_full", classify_selection_execution_html("alert('总容量已满，请选择其它课程！')")[0])
        self.assertEqual("rejected", classify_selection_execution_html("alert('与已选课程时间冲突')")[0])
        self.assertEqual("unknown", classify_selection_execution_html("<html>no result</html>")[0])

    def test_adapter_reloads_exact_section_and_submits_once(self):
        class Page:
            url = "https://webvpn.hitwh.edu.cn/proxy/xsxk/queryXsxkList"
            evaluations = 0

            def evaluate(self, script, payload):
                self.evaluations += 1
                if self.evaluations == 1:
                    self.query_payload = payload
                    return {"status": 200, "body": COURSE_HTML}
                self.payload = payload
                return {"status": 200, "body": "alert('选课成功')"}

        page = Page()
        opened = []
        result = VerifiedSelectionExecutionAdapter().execute(
            page,
            section_id="TERM-TASK-001",
            category="xsxk",
            term_value="2026-20271",
            authenticate=lambda url, target: opened.append(url) or target,
        )
        self.assertEqual("selected", result.status)
        self.assertEqual(2, page.evaluations)
        self.assertIn("pageXklb=xsxk", opened[0])
        self.assertEqual("", page.query_payload["overrides"]["rwh"])
        self.assertIn("pageNo", page.query_payload["remove"])
        self.assertEqual("TERM-TASK-001", page.payload["overrides"]["rwh"])

    def test_adapter_does_not_submit_missing_section(self):
        class Page:
            url = "https://webvpn.hitwh.edu.cn/proxy/xsxk/queryXsxkList"
            def evaluate(self, *_args):
                return {"status": 200, "body": COURSE_HTML}

        with self.assertRaisesRegex(RuntimeError, "已不在"):
            VerifiedSelectionExecutionAdapter().execute(
                Page(), section_id="OTHER", category="xsxk", term_value="2026-20271",
                authenticate=lambda _url, target: target,
            )


class ExecutionPreparationTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.database = WorkspaceDatabase.open(Path(self.directory.name))
        with self.database.connection:
            profile_id = self.database._insert_profile({"grade": "2025"})
            notice_id = self.database._insert_notice({
                "status": "confirmed", "term": "2026-2027-1", "windows": [{
                    "action": "selection", "method": "academic_system", "grades": ["2025"],
                    "category_codes": ["xsxk"],
                }],
            })
            self.database.connection.execute("update notice_versions set status='confirmed' where id=?", (notice_id,))
        self.profile_id, self.notice_id = profile_id, notice_id
        self.timetable = self.database.publish_snapshot(
            "timetable", "2026-2027-1", {"status": "complete", "entries": []},
            source="test", profile_id=profile_id,
        )
        self.selection = self.database.publish_snapshot(
            "selection", "2026-2027-1", {"status": "complete", "sections": [{
                "identity": "TERM-TASK-001", "action_rwh": "TERM-TASK-001",
                "execution_ready": True, "course_code": "C1", "name": "课程",
                "query_code": "xsxk", "query_term": "2026-20271", "query_page": 3,
                "meetings": [{"day": 1, "start": 1, "end": 2, "weeks": [1]}],
            }]}, source="test", profile_id=profile_id, notice_id=notice_id,
        )
        self.service = WorkbenchService(self.database)

    def tearDown(self):
        self.database.close()
        self.directory.cleanup()

    def test_maps_exact_current_snapshot_section(self):
        context = self.service.prepare_selection_execution("TERM-TASK-001", self.selection["id"])
        self.assertEqual("TERM-TASK-001", context["section_id"])
        self.assertEqual("xsxk", context["category"])
        self.assertEqual(3, context["source_page"])

    def test_stale_or_missing_snapshot_is_blocked(self):
        with self.assertRaisesRegex(ValueError, "快照已更新"):
            self.service.prepare_selection_execution("TERM-TASK-001", "old")
        with self.assertRaisesRegex(ValueError, "不存在"):
            self.service.prepare_selection_execution("OTHER", self.selection["id"])

    def test_conflict_and_unknown_schedule_are_blocked(self):
        self.database.publish_snapshot(
            "timetable", "2026-2027-1", {"status": "complete", "entries": [{
                "course_code": "BUSY", "meetings": [{"day": 1, "start": 2, "end": 3, "weeks": [1]}],
            }]}, source="test", profile_id=self.profile_id,
        )
        with self.assertRaisesRegex(ValueError, "current_timetable"):
            self.service.prepare_selection_execution("TERM-TASK-001", self.selection["id"])

        self.database.publish_snapshot(
            "timetable", "2026-2027-1", {"status": "complete", "entries": [{"course_code": "UNKNOWN"}]},
            source="test", profile_id=self.profile_id,
        )
        with self.assertRaisesRegex(ValueError, "conflict_unknown"):
            self.service.prepare_selection_execution("TERM-TASK-001", self.selection["id"])


class ExecutionApiTests(unittest.TestCase):
    def test_csrf_bound_confirmation_and_single_execution_task(self):
        class Gateway(FakeGateway):
            def __init__(self):
                super().__init__()
                self.execution_count = 0

            def execute_selection(self, context, progress, cancelled):
                self.execution_count += 1
                return {"status": "selected", "message": "选课成功", "section_id": context["section_id"]}

        with tempfile.TemporaryDirectory() as directory:
            gateway = Gateway()
            app = create_workbench_app(Path(directory), gateway_factory=lambda: gateway)
            database = app.extensions["workspace_database"]
            with database.connection:
                profile_id = database._insert_profile({"grade": "2025"})
                notice_id = database._insert_notice({
                    "status": "confirmed", "term": "2026-2027-1", "windows": [{
                        "action": "selection", "method": "academic_system", "grades": ["2025"],
                        "category_codes": ["xsxk"],
                    }],
                })
                database.connection.execute("update notice_versions set status='confirmed' where id=?", (notice_id,))
            database.publish_snapshot("timetable", "2026-2027-1", {"status": "complete", "entries": []}, source="test", profile_id=profile_id)
            selection = database.publish_snapshot("selection", "2026-2027-1", {"status": "complete", "sections": [{
                "identity": "TERM-TASK-001", "action_rwh": "TERM-TASK-001", "execution_ready": True,
                "course_code": "C1", "name": "课程", "query_code": "xsxk", "query_term": "2026-20271",
                "meetings": [{"day": 1, "start": 1, "end": 2, "weeks": [1]}],
            }]}, source="test", profile_id=profile_id, notice_id=notice_id)
            client = app.test_client()
            body = {"section_id": "TERM-TASK-001", "snapshot_id": selection["id"], "confirmation": "TERM-TASK-001"}
            self.assertEqual(403, client.post("/api/executions/selection", json=body).status_code)
            state = client.get("/api/state").get_json()
            headers = {"Origin": "http://localhost", "Host": "localhost", "X-CSRF-Token": state["csrf_token"]}
            rejected = client.post("/api/executions/selection", json={**body, "confirmation": True}, headers=headers)
            self.assertEqual(409, rejected.status_code)
            created = client.post("/api/executions/selection", json=body, headers=headers)
            self.assertEqual(202, created.status_code)
            task = created.get_json()
            self.assertEqual("execution", task["task_kind"])
            service = app.extensions["observation_service"]
            self.assertTrue(service.wait(task["id"], 2))
            inspected = client.get(f"/api/tasks/{task['id']}").get_json()
            self.assertEqual("selected", inspected["progress"]["result"]["status"])
            history = client.get("/api/executions").get_json()["executions"]
            self.assertEqual("selected", history[0]["result"])
            self.assertEqual("TERM-TASK-001", history[0]["section_id"])
            self.assertNotIn("token", json.dumps(history).lower())
            self.assertEqual(1, gateway.execution_count)
            service.close()
            database.close()



class UnknownExecutionTests(unittest.TestCase):
    def test_unknown_result_blocks_until_user_resolves_it(self):
        with tempfile.TemporaryDirectory() as directory:
            database = WorkspaceDatabase.open(Path(directory))
            database.record_execution(
                {"section_id": "TASK", "course_name": "课程", "category": "xsxk", "snapshot_id": "S", "notice_id": "N"},
                {"status": "unknown", "message": "未识别到选课结果"},
            )
            blocked = database.unresolved_execution("TASK")
            self.assertIsNotNone(blocked)
            self.assertTrue(database.resolve_execution(blocked["id"]))
            self.assertIsNone(database.unresolved_execution("TASK"))
            database.close()


if __name__ == "__main__":
    unittest.main()
