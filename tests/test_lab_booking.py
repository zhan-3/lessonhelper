import unittest

from course_selection.lab_booking import (
    LabSlot,
    book_slots,
    busy_from_bookings,
    busy_intervals,
    plan_lab_slots,
    plan_token,
    same_time_pairs,
    week_numbers,
)
from course_selection.lab_browser_session import BrowserLabSession


def entry(weekday, start, end, weeks, label="", parity="all"):
    return {
        "weekday": weekday,
        "start_period": start,
        "end_period": end,
        "week_start": weeks[0],
        "week_end": weeks[1],
        "week_parity": parity,
        "course_name": label,
        "conflict_status": "ok",
    }


class ScheduleRow(dict):
    pass


def row(class_date, week, weekday_name, timers):
    return {
        "classDate": class_date,
        "eduWeek": week,
        "dayWeek": weekday_name,
        "timerList": [{"timer": item[0], "timerName": item[1], "startTime": item[2]} for item in timers],
    }


class FakeSession:
    """Recording fake of the center app surface."""

    def __init__(self, subjects, schedules, seats, booked=()):
        self._subjects = subjects
        self._schedules = schedules
        self._seats = seats
        self._booked = list(booked)
        self.submitted = []
        self.seat_queries = []
        self.submit_response = {"transport": "ok", "code": 0, "message": "请求成功", "result": True}

    def booked(self):
        return self._booked

    def subjects(self):
        return self._subjects

    def schedule(self, subject_id):
        return self._schedules.get(subject_id, [])

    def free_seat(self, subject_id, class_date, timer):
        self.seat_queries.append((subject_id, class_date, timer))
        return self._seats.get((subject_id, class_date, timer))

    def submit(self, slot):
        self.submitted.append(slot)
        return self.submit_response


SUBJECTS = [{"subjectId": 3001, "subjectName": "电表改装与校准"},
            {"subjectId": 3002, "subjectName": "动力学基础实验"}]

SEATS = {
    (3001, "2026-09-22", 2): {"seatId": "1", "tableNo": "1", "lab": "N544"},
    (3002, "2026-09-22", 2): {"seatId": "9", "tableNo": "9", "lab": "N535"},
    (3002, "2026-09-23", 2): {"seatId": "7", "tableNo": "7", "lab": "N535"},
}


class WeekNumberTests(unittest.TestCase):
    def test_parity_and_explicit_lists_are_normalized(self):
        self.assertEqual(frozenset({1, 3, 5}), week_numbers(entry(1, 1, 2, (1, 5), parity="odd")))
        self.assertEqual(frozenset({2, 4}), week_numbers(entry(1, 1, 2, (1, 5), parity="even")))
        self.assertEqual(frozenset({4, 6}), week_numbers({"week_numbers": [4, 6]}))


class BusyIntervalTests(unittest.TestCase):
    def test_unknown_time_rows_are_excluded(self):
        entries = [entry(1, 1, 2, (1, 16)), {**entry(2, 1, 2, (1, 16)), "conflict_status": "unknown"}]
        self.assertEqual(1, len(busy_intervals(entries)))

    def test_overlap_uses_weekday_week_and_periods(self):
        (interval,) = busy_intervals([entry(2, 3, 4, (1, 16), "专业课")])
        self.assertTrue(interval.overlaps(2, 5, (3, 4)))
        self.assertTrue(interval.overlaps(2, 5, (2, 3)))
        self.assertFalse(interval.overlaps(3, 5, (3, 4)))
        self.assertFalse(interval.overlaps(2, 5, (5, 6)))


class BookingIntervalTests(unittest.TestCase):
    def test_booked_lab_sessions_become_occupied_intervals(self):
        (interval,) = busy_from_bookings([
            {"classDate": "2026-09-22", "eduWeek": 4, "startTime": "10:05:00", "subjectName": "惠斯通电桥测电阻"},
        ])
        self.assertEqual(2, interval.weekday)
        self.assertEqual((3, 4), (interval.start_period, interval.end_period))
        self.assertEqual(frozenset({4}), interval.weeks)
        self.assertTrue(interval.overlaps(2, 4, (3, 4)))
        self.assertFalse(interval.overlaps(2, 5, (3, 4)))

    def test_rows_without_a_known_start_time_are_ignored(self):
        self.assertEqual((), busy_from_bookings([{"classDate": "2026-09-22", "eduWeek": 4, "startTime": "11:11:00"}]))

    def test_planner_avoids_slots_already_taken_by_another_course(self):
        schedules = {3002: [row("2026-09-22", 4, "星期二", [(2, "第二大节", "10:05:00")]),
                            row("2026-09-23", 4, "星期三", [(2, "第二大节", "10:05:00")])]}
        session = FakeSession([SUBJECTS[1]], schedules, SEATS)
        busy = busy_from_bookings([
            {"classDate": "2026-09-22", "eduWeek": 4, "startTime": "10:05:00", "subjectName": "已约实验"},
        ])
        result = plan_lab_slots(session, busy)
        self.assertEqual("2026-09-23", result.slots[0].class_date)


class PlanTests(unittest.TestCase):
    def test_planner_skips_timetable_clashes(self):
        schedules = {
            3001: [row("2026-09-22", 4, "星期二", [(2, "第二大节", "10:05:00")])],
            3002: [
                row("2026-09-22", 4, "星期二", [(2, "第二大节", "10:05:00")]),
                row("2026-09-23", 4, "星期三", [(2, "第二大节", "10:05:00")]),
            ],
        }
        session = FakeSession(SUBJECTS, schedules, SEATS)
        busy = busy_intervals([entry(2, 3, 4, (1, 16), "周二专业课")])

        result = plan_lab_slots(session, busy)

        self.assertEqual([3002], [slot.subject_id for slot in result.slots])
        self.assertEqual("2026-09-23", result.slots[0].class_date)
        self.assertEqual(["电表改装与校准"], result.unresolved)
        self.assertEqual([], result.reminders)

    def test_planner_reports_same_time_pairs_as_reminders(self):
        schedules = {
            3001: [row("2026-09-22", 4, "星期二", [(2, "第二大节", "10:05:00")])],
            3002: [row("2026-09-22", 4, "星期二", [(2, "第二大节", "10:05:00")])],
        }
        session = FakeSession(SUBJECTS, schedules, SEATS)
        result = plan_lab_slots(session, busy_intervals([entry(3, 3, 4, (1, 16), "周三专业课")]))

        self.assertEqual([3001, 3002], [slot.subject_id for slot in result.slots])
        self.assertEqual([], result.unresolved)
        self.assertEqual(1, len(result.reminders))
        left, right = result.reminders[0]
        self.assertEqual("2026-09-22", left.class_date)
        self.assertEqual(left.timer, right.timer)

    def test_already_booked_subjects_are_not_planned_again(self):
        schedules = {3002: [row("2026-09-23", 4, "星期三", [(2, "第二大节", "10:05:00")])]}
        session = FakeSession(SUBJECTS, schedules, SEATS, booked=[{"subjectId": 3001}])
        result = plan_lab_slots(session, ())
        self.assertEqual([3002], [slot.subject_id for slot in result.slots])

    def test_unresolved_experiments_are_named(self):
        session = FakeSession(SUBJECTS, {3001: [], 3002: []}, {})
        result = plan_lab_slots(session, ())
        self.assertEqual([], result.slots)
        self.assertEqual(["电表改装与校准", "动力学基础实验"], result.unresolved)

    def test_probe_cap_bounds_availability_queries(self):
        schedules = {
            3001: [row(f"2026-09-{day:02d}", 4, "星期二", [(2, "第二大节", "10:05:00")])
                   for day in range(21, 30)]
        }
        session = FakeSession([SUBJECTS[0]], schedules, {})
        plan_lab_slots(session, (), probe_cap=3)
        self.assertEqual(3, len(session.seat_queries))


class BookingTests(unittest.TestCase):
    def setUp(self):
        self.schedules = {3001: [row("2026-09-22", 4, "星期二", [(2, "第二大节", "10:05:00")])]}
        self.session = FakeSession([SUBJECTS[0]], self.schedules, SEATS)
        self.slots = plan_lab_slots(self.session, ()).slots

    def test_confirmation_token_must_match_the_exact_targets(self):
        with self.assertRaises(ValueError):
            book_slots(self.session, self.slots, confirmation="whatever")
        self.assertEqual([], self.session.submitted)

    def test_successful_submission_happens_once(self):
        outcomes = book_slots(self.session, self.slots, confirmation=plan_token(self.slots))
        self.assertEqual(["confirmed_success"], [item["outcome"] for item in outcomes])
        self.assertEqual(1, len(self.session.submitted))

    def test_failure_stops_the_run_without_retrying(self):
        self.session.submit_response = {"transport": "ok", "code": 5000, "message": "此座位已被预约，请选择其它座位"}
        outcomes = book_slots(self.session, self.slots, confirmation=plan_token(self.slots))
        self.assertEqual(["confirmed_failure"], [item["outcome"] for item in outcomes])
        self.assertEqual(1, len(self.session.submitted))

    def test_transport_error_is_recorded_as_possibly_applied(self):
        self.session.submit_response = {"transport": "error", "detail": "timeout"}
        outcomes = book_slots(self.session, self.slots, confirmation=plan_token(self.slots))
        self.assertEqual(["possibly_applied"], [item["outcome"] for item in outcomes])
        self.assertEqual(1, len(self.session.submitted))

    def test_seat_taken_before_submit_skips_without_submitting(self):
        self.session._seats = {}
        outcomes = book_slots(self.session, self.slots, confirmation=plan_token(self.slots))
        self.assertEqual(["skipped_unavailable"], [item["outcome"] for item in outcomes])
        self.assertEqual([], self.session.submitted)

    def test_same_time_pairs_are_reported_not_hidden(self):
        pairs = same_time_pairs([
            LabSlot(1, "a", "2026-09-22", 4, 2, 2, "第二大节", "10:05", "N1", "1", "1"),
            LabSlot(2, "b", "2026-09-22", 4, 2, 2, "第二大节", "10:05", "N2", "1", "2"),
            LabSlot(3, "c", "2026-09-23", 4, 3, 2, "第二大节", "10:05", "N3", "1", "3"),
        ])
        self.assertEqual(1, len(pairs))


class AdapterTests(unittest.TestCase):
    def test_adapter_does_not_persist_the_token_and_uses_the_write_endpoint_once(self):
        calls = []

        class Page:
            def evaluate(self, script, argument=None):
                if argument is None:
                    return "token-value"
                calls.append(argument)
                if "doyyxkzw" in argument[0]:
                    return {"transport": "ok", "code": 0, "message": "请求成功", "result": True}
                return {"transport": "ok", "code": 0, "result": {"data": {"labList": [], "lab": {}}}}

        session = BrowserLabSession(Page(), "dxwl", "token-value", pause=0)
        slot = LabSlot(3001, "电表改装与校准", "2026-09-22", 4, 2, 2, "第二大节", "10:05", "N544", "1", "1")
        payload = session.submit(slot)

        self.assertTrue(payload["result"])
        path, form, token = calls[-1]
        self.assertIn("doyyxkzw", path)
        self.assertEqual({"subjectId": 3001, "classDate": "2026-09-22", "timer": "10:05:00",
                          "timercode": 2, "id": "1"}, form)
        self.assertEqual("token-value", token)


if __name__ == "__main__":
    unittest.main()
