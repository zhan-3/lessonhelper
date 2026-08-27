import json
import os
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path

from course_selection.cli import build_parser
from course_selection.student_profile_observation import (
    _validated_academic_entry,
    cleanup_expired_analyses,
    profile_candidate_from_payload,
)


class StudentProfileObservationTests(unittest.TestCase):
    def test_cli_exposes_only_the_student_profile_target(self):
        args = build_parser().parse_args(
            ["analyze-interface", "--target", "student-profile", "--wait-seconds", "10"]
        )
        self.assertEqual("analyze-interface", args.command)
        self.assertEqual("student-profile", args.target)
        self.assertEqual(10, args.wait_seconds)
        with self.assertRaises(SystemExit):
            build_parser().parse_args(["analyze-interface", "--target", "selection"])

    def test_candidate_keeps_field_structure_without_identity_values(self):
        payload = {
            "data": {
                "xm": "测试姓名",
                "xh": "2025000000",
                "zymc": "示例专业",
                "rxnj": "2025",
            }
        }
        candidate = profile_candidate_from_payload(
            url="https://academic.example/api/profile?token=secret",
            method="GET",
            status=200,
            payload=payload,
        )
        serialized = json.dumps(candidate, ensure_ascii=False)
        self.assertGreater(candidate["score"], 10)
        self.assertEqual(
            {"grade", "major", "student_name", "student_number"},
            set(candidate["matched_fields"]),
        )
        self.assertNotIn("测试姓名", serialized)
        self.assertNotIn("2025000000", serialized)
        self.assertNotIn("示例专业", serialized)
        self.assertNotIn("secret", serialized)

    def test_identity_values_used_as_keys_and_url_fragments_are_redacted(self):
        candidate = profile_candidate_from_payload(
            url="https://academic.example/profile#ticket=secret-value",
            method="GET",
            status=200,
            payload={"2025000000": {"xm": "测试姓名", "张三": {"major": "示例"}}},
        )
        serialized = json.dumps(candidate, ensure_ascii=False)
        self.assertNotIn("2025000000", serialized)
        self.assertNotIn("张三", serialized)
        self.assertNotIn("secret-value", serialized)
        self.assertIn("[dynamic-key]", serialized)

    def test_unrelated_application_resource_is_not_a_profile_candidate(self):
        candidate = profile_candidate_from_payload(
            url="https://example.test/resources",
            method="GET",
            status=200,
            payload={"data": [{"resource": [{"name": "新教务系统"}]}]},
        )
        self.assertIsNone(candidate)

    def test_academic_entry_must_be_an_exact_webvpn_proxy_path(self):
        current = "https://webvpn.hitwh.edu.cn/portal/#!/service"
        valid = _validated_academic_entry(current, "/http/abcdef1234/")
        self.assertEqual("https://webvpn.hitwh.edu.cn/http/abcdef1234/", valid)
        with self.assertRaises(ValueError):
            _validated_academic_entry(current, "https://attacker.example/academic")
        with self.assertRaises(ValueError):
            _validated_academic_entry(current, "/unexpected/path")

    def test_cleanup_removes_only_expired_analysis_directories(self):
        now = datetime(2026, 8, 27, tzinfo=timezone.utc)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            old = root / "old"
            recent = root / "recent"
            old.mkdir()
            recent.mkdir()
            old_time = (now - timedelta(days=8)).timestamp()
            recent_time = (now - timedelta(days=2)).timestamp()
            os.utime(old, (old_time, old_time))
            os.utime(recent, (recent_time, recent_time))

            removed = cleanup_expired_analyses(root, now=now, retention_days=7)

            self.assertEqual(1, removed)
            self.assertFalse(old.exists())
            self.assertTrue(recent.exists())


if __name__ == "__main__":
    unittest.main()
