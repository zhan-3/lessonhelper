import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from click.testing import CliRunner

from course_selection.cli import main
from course_selection.lab_contract import ContractSnapshot

SUBJECTS = {"transport": "ok", "code": 0, "result": [
    {"subjectId": 3001, "subjectName": "电表改装与校准", "izPass": False}]}
SEATS = {"transport": "ok", "code": 0, "result": {"msg": "", "data": {
    "labList": [{"TableNo": "1", "id": "21989887", "status": "空闲"}],
    "lab": {"LabsName": "N544"}}}}


class FakeTransport:
    """Read-only stand-in for one center app."""

    center = "dxwl"

    def __init__(self, *, drop_subject_field=False, fail_subjects=False):
        self.drop_subject_field = drop_subject_field
        self.fail_subjects = fail_subjects
        self.closed = False

    def call(self, path, form=None):
        if path == "auth/currentUserInfo":
            return {"transport": "ok", "code": 0, "result": {"userId": "1"}}
        if path == "view/lesson/ckkb":
            return {"transport": "ok", "code": 0, "result": []}
        if path == "view/subjects":
            if self.fail_subjects:
                return {"transport": "error", "detail": "OSError: no route"}
            subjects = [dict(row) for row in SUBJECTS["result"]]
            if self.drop_subject_field:
                for row in subjects:
                    row.pop("izPass", None)
            return {"transport": "ok", "code": 0, "result": subjects}
        if path == "view/booking/yyxh":
            return {"transport": "ok", "code": 0, "result": {"data": [
                {"classDate": "2026-09-22", "eduWeek": 4, "dayWeek": "星期二",
                 "timerList": [{"timer": 2, "timerName": "第二大节", "startTime": "10:05:00"}]}]}}
        if path == "view/booking/yyxkzw":
            return SEATS
        return {"transport": "error", "detail": "not stubbed"}

    def close(self):
        self.closed = True


class LabContractCliTests(unittest.TestCase):
    def setUp(self):
        self.runner = CliRunner()
        self.tmp = tempfile.TemporaryDirectory()
        root = Path(self.tmp.name)
        self.observations = root / "obs"
        self.baselines = root / "base"

    def tearDown(self):
        self.tmp.cleanup()

    def run_cli(self, action, transport, *extra):
        with mock.patch("course_selection.cli._lab_contract_transport", return_value=transport):
            return self.runner.invoke(main, [
                "lab-contract", action,
                "--observations-dir", str(self.observations),
                "--baselines-dir", str(self.baselines),
                *extra,
            ])

    def test_record_writes_an_observation_and_says_there_is_no_baseline_yet(self):
        transport = FakeTransport()
        result = self.run_cli("record", transport)

        self.assertEqual(0, result.exit_code, result.output)
        self.assertTrue(transport.closed)
        observed = json.loads((self.observations / "direct-dxwl.observed.json").read_text(encoding="utf-8"))
        self.assertIn("view/subjects", observed["endpoints"])
        self.assertEqual("direct", observed["environment"]["channel"])
        self.assertIn("尚无基线", result.output)

    def test_check_without_a_baseline_cannot_evaluate(self):
        result = self.run_cli("check", FakeTransport())
        self.assertEqual(2, result.exit_code, result.output)
        self.assertIn("无法比对", result.output)

    def test_promote_strips_the_machine_local_environment(self):
        self.run_cli("record", FakeTransport())
        promoted = self.run_cli("promote", FakeTransport())

        self.assertEqual(0, promoted.exit_code, promoted.output)
        baseline = ContractSnapshot.from_json((self.baselines / "direct-dxwl.json").read_text(encoding="utf-8"))
        self.assertEqual({}, baseline.environment)
        self.assertTrue(baseline.endpoints)

    def test_check_reports_a_clean_baseline_as_no_change(self):
        self.run_cli("record", FakeTransport())
        self.run_cli("promote", FakeTransport())

        result = self.run_cli("check", FakeTransport())
        self.assertEqual(0, result.exit_code, result.output)
        self.assertIn("无变化", result.output)

    def test_check_blocks_when_a_locked_field_disappears(self):
        self.run_cli("record", FakeTransport())
        self.run_cli("promote", FakeTransport())
        baseline_path = self.baselines / "direct-dxwl.json"
        payload = json.loads(baseline_path.read_text(encoding="utf-8"))
        payload["endpoints"]["view/subjects"]["locked"] = ["izPass"]
        baseline_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

        result = self.run_cli("check", FakeTransport(drop_subject_field=True))

        self.assertEqual(1, result.exit_code, result.output)
        self.assertIn("breaking", result.output)

    def test_check_blocks_when_an_endpoint_cannot_be_read(self):
        self.run_cli("record", FakeTransport())
        self.run_cli("promote", FakeTransport())

        result = self.run_cli("check", FakeTransport(fail_subjects=True))

        self.assertEqual(1, result.exit_code, result.output)
        self.assertIn("unavailable", result.output)

    def test_json_output_is_machine_readable(self):
        self.run_cli("record", FakeTransport())
        self.run_cli("promote", FakeTransport())

        result = self.run_cli("check", FakeTransport(), "--json")

        self.assertEqual(0, result.exit_code, result.output)
        self.assertEqual("dxwl", json.loads(result.output)["center"])

    def test_promote_keeps_previously_locked_fields(self):
        self.run_cli("record", FakeTransport())
        self.run_cli("promote", FakeTransport())
        baseline_path = self.baselines / "direct-dxwl.json"
        payload = json.loads(baseline_path.read_text(encoding="utf-8"))
        payload["endpoints"]["view/subjects"]["locked"] = ["izPass"]
        baseline_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

        self.run_cli("record", FakeTransport())
        self.run_cli("promote", FakeTransport())

        kept = ContractSnapshot.from_json(baseline_path.read_text(encoding="utf-8"))
        self.assertEqual(("izPass",), kept.endpoints["view/subjects"].locked)


if __name__ == "__main__":
    unittest.main()
