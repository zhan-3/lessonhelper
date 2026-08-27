import unittest
from datetime import datetime, timezone, timedelta

from course_selection.planning import build_read_only_plan


NOW = datetime(2026, 8, 26, tzinfo=timezone.utc)


def snapshot(kind, identity, *, profile="profile-1", notice=None, term="2026-1", payload=None, age=0):
    return {
        "id": identity, "kind": kind, "term": term, "profile_id": profile,
        "notice_id": notice, "source_at": (NOW - timedelta(seconds=age)).isoformat(),
        "payload": payload or {"status": "complete"},
    }


class PlanningTests(unittest.TestCase):
    def setUp(self):
        self.timetable = snapshot("timetable", "tt-1", payload={"status": "complete", "entries": [
            {"course_code": "MATH", "course_name": "Math", "weekday": 1, "start_period": 1, "end_period": 2, "week_numbers": [1, 2]},
        ]})
        self.selection = snapshot("selection", "sel-1", notice="notice-1", payload={"status": "complete", "sections": [
            {"identity": "free", "name": "Free", "weekday": 2, "start_period": 3, "end_period": 4, "week_numbers": [1, 2]},
            {"identity": "busy", "name": "Busy", "weekday": 1, "start_period": 2, "end_period": 3, "week_numbers": [1, 2]},
            {"identity": "unknown", "name": "Pending", "time": "pending"},
        ]})

    def test_builds_ranked_goals_and_detects_current_and_candidate_conflicts(self):
        plan = build_read_only_plan(
            term="2026-1", profile_id="profile-1", notice_id="notice-1",
            timetable_snapshot=self.timetable, selection_snapshot=self.selection,
            goals=[{"goal_id": "g1", "course_identity": "COURSE", "rank": 1,
                     "preferences": [{"section_id": "busy", "rank": 1}, {"section_id": "free", "rank": 2}]}], now=NOW)
        self.assertTrue(plan.ready)
        self.assertEqual(["busy", "free"], [p.section_id for p in plan.goals[0].preferences])
        self.assertTrue(any(c.kind == "current_timetable" for c in plan.conflicts))

    def test_unknown_schedule_is_not_treated_as_free(self):
        plan = build_read_only_plan(
            term="2026-1", profile_id="profile-1", notice_id="notice-1",
            timetable_snapshot=self.timetable, selection_snapshot=self.selection,
            goals=[{"course_identity": "COURSE", "preferences": [{"section_id": "unknown"}]}], now=NOW)
        self.assertTrue(any(c.kind == "conflict_unknown" for c in plan.conflicts))
        self.assertFalse(plan.ready)
        self.assertIn("conflict_unknown", plan.blocked_reasons)

    def test_stale_or_mismatched_snapshots_block_plan(self):
        stale = snapshot("selection", "sel-old", notice="old-notice", age=1801,
                         payload={"status": "complete", "sections": []})
        plan = build_read_only_plan(
            term="2026-1", profile_id="profile-1", notice_id="notice-1",
            timetable_snapshot=self.timetable, selection_snapshot=stale,
            goals=[{"course_identity": "COURSE"}], now=NOW)
        self.assertFalse(plan.ready)
        self.assertIn("selection_notice_mismatch", plan.blocked_reasons)
        self.assertIn("selection_snapshot_stale", plan.blocked_reasons)

    def test_missing_section_blocks_plan_without_academic_request(self):
        plan = build_read_only_plan(
            term="2026-1", profile_id="profile-1", notice_id="notice-1",
            timetable_snapshot=self.timetable, selection_snapshot=self.selection,
            goals=[{"course_identity": "COURSE", "preferences": [{"section_id": "gone"}]}], now=NOW)
        self.assertFalse(plan.ready)
        self.assertIn("section_missing:gone", plan.blocked_reasons)


if __name__ == "__main__":
    unittest.main()
