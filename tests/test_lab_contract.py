import json
import unittest

from course_selection.lab_contract import (
    ADDITIVE,
    BREAKING,
    COSMETIC,
    NOTICE,
    REBASELINE,
    UNAVAILABLE,
    ContractSnapshot,
    EndpointContract,
    diff,
    fingerprint,
    observe,
    with_locked,
)


def snapshot(**overrides):
    base = {
        "center": "dxwl",
        "channel": "direct",
        "origin": "http://openlab.hitwh.edu.cn",
        "app_version": "2.8.0",
        "environment": {"proxy": "127.0.0.1:7897", "dns_source": "clash_doh"},
        "endpoints": {
            "view/subjects": EndpointContract(
                "view/subjects",
                response_fields=("subjectId", "subjectName", "izPass"),
                locked=("subjectId", "izPass"),
                envelope=("code", "message", "result", "timestamp"),
                codes=("0",),
            )
        },
    }
    base.update(overrides)
    return ContractSnapshot(**base)


def severities(report):
    return {(item.severity, item.subject) for item in report.items}


class FingerprintTests(unittest.TestCase):
    def test_fingerprint_is_stable_and_ignores_field_order(self):
        first = snapshot()
        second = snapshot(endpoints={
            "view/subjects": EndpointContract("view/subjects",
                                              response_fields=("izPass", "subjectId", "subjectName"),
                                              locked=("izPass", "subjectId"))
        })
        self.assertEqual(fingerprint(first), fingerprint(second))

    def test_fingerprint_changes_when_a_contract_field_disappears(self):
        changed = snapshot(endpoints={
            "view/subjects": EndpointContract("view/subjects", response_fields=("subjectId",))
        })
        self.assertNotEqual(fingerprint(snapshot()), fingerprint(changed))


class SerializationTests(unittest.TestCase):
    def test_round_trip_keeps_every_recorded_fact(self):
        restored = ContractSnapshot.from_json(snapshot().to_json())
        self.assertEqual(snapshot().to_dict(), restored.to_dict())

    def test_unknown_schema_version_is_rejected(self):
        with self.assertRaises(ValueError):
            ContractSnapshot.from_dict({"schema_version": 99, "center": "dxwl"})


class DiffTests(unittest.TestCase):
    def test_identical_snapshots_report_no_change(self):
        report = diff(snapshot(), snapshot())
        self.assertEqual([], report.items)
        self.assertFalse(report.blocks)
        self.assertEqual(["  无变化"], report.describe())

    def test_version_and_asset_churn_is_cosmetic_only(self):
        current = snapshot(app_version="2.8.1", static_assets=("index.abc.js",))
        report = diff(snapshot(), current)
        self.assertEqual(2, report.count(COSMETIC))
        self.assertFalse(report.blocks)

    def test_channel_or_origin_change_asks_for_a_new_baseline(self):
        report = diff(snapshot(), snapshot(channel="webvpn"))
        self.assertIn((REBASELINE, "channel"), severities(report))
        report = diff(snapshot(), snapshot(origin="https://webvpn.hitwh.edu.cn"))
        self.assertIn((REBASELINE, "origin"), severities(report))

    def test_environment_drift_is_only_a_notice(self):
        current = snapshot(environment={"proxy": "off", "dns_source": "campus"})
        report = diff(snapshot(), current)
        self.assertEqual(2, report.count(NOTICE))
        self.assertFalse(report.blocks)

    def test_removed_endpoint_blocks(self):
        report = diff(snapshot(), snapshot(endpoints={}))
        self.assertIn((BREAKING, "view/subjects"), severities(report))
        self.assertTrue(report.blocks)

    def test_unreadable_endpoint_blocks_without_claiming_it_was_removed(self):
        report = diff(snapshot(), snapshot(endpoints={}, unavailable=("view/subjects",)))
        self.assertIn((UNAVAILABLE, "view/subjects"), severities(report))
        self.assertNotIn((BREAKING, "view/subjects"), severities(report))
        self.assertTrue(report.blocks)

    def test_locked_field_loss_blocks_but_unlocked_field_loss_does_not(self):
        locked_loss = snapshot(endpoints={
            "view/subjects": EndpointContract("view/subjects", response_fields=("subjectName", "izPass"),
                                              locked=("subjectId", "izPass"))
        })
        report = diff(snapshot(), locked_loss)
        self.assertIn((BREAKING, "view/subjects.subjectId"), severities(report))
        self.assertTrue(report.blocks)

        unlocked_loss = snapshot(endpoints={
            "view/subjects": EndpointContract("view/subjects", response_fields=("subjectId", "izPass"),
                                              locked=("subjectId", "izPass"))
        })
        report = diff(snapshot(), unlocked_loss)
        self.assertIn((ADDITIVE, "view/subjects.subjectName"), severities(report))
        self.assertFalse(report.blocks)

    def test_new_field_and_new_code_are_additive(self):
        current = snapshot(endpoints={
            "view/subjects": EndpointContract("view/subjects",
                                              response_fields=("subjectId", "subjectName", "izPass", "izNow"),
                                              locked=("subjectId", "izPass"), codes=("0", "5000"))
        })
        report = diff(snapshot(), current)
        self.assertIn((ADDITIVE, "view/subjects.izNow"), severities(report))
        self.assertIn((ADDITIVE, "view/subjects.code[5000]"), severities(report))
        self.assertFalse(report.blocks)

    def test_endpoint_without_a_claim_on_either_side_is_not_drift(self):
        empty = {"view/subjects": EndpointContract("view/subjects", response_fields=())}
        report = diff(snapshot(endpoints=empty), snapshot(endpoints=empty))
        self.assertEqual([], report.items)
        self.assertFalse(report.blocks)

    def test_unreadable_endpoint_is_unavailable_rather_than_unchanged(self):
        current = snapshot(endpoints={
            "view/subjects": EndpointContract("view/subjects", response_fields=(), locked=("subjectId",))
        })
        report = diff(snapshot(), current)
        self.assertIn((UNAVAILABLE, "view/subjects"), severities(report))
        self.assertTrue(report.blocks)
        self.assertEqual({"blocks": True, "center": "dxwl", "items": report.to_dict()["items"]},
                         json.loads(json.dumps(report.to_dict())))


class FakeTransport:
    """Read-only fake of a center app, recording the forms it was asked for."""

    center = "dxwl"

    def __init__(self, responses):
        self.responses = responses
        self.calls = []

    def call(self, path, form=None):
        self.calls.append((path, dict(form or {})))
        return self.responses.get(path, {"transport": "error", "detail": "not stubbed"})

    def close(self):
        pass


SUBJECTS = {"transport": "ok", "code": 0, "message": "请求成功", "timestamp": "t",
            "result": [{"subjectId": 3001, "subjectName": "电表改装与校准", "izPass": False}]}
SCHEDULE = {"transport": "ok", "code": 0, "result": {"count": 1, "data": [
    {"classDate": "2026-09-22", "eduWeek": 4, "dayWeek": "星期二",
     "timerList": [{"timer": 2, "timerName": "第二大节", "startTime": "10:05:00"}]}]}}
SEATS = {"transport": "ok", "code": 0, "result": {"msg": "", "data": {
    "labList": [{"TableShow": "1", "TableNo": "1", "id": "21989887", "status": "空闲"}],
    "lab": {"LabsName": "N544", "LabsAddress": "N544"}}}}


class ObserveTests(unittest.TestCase):
    def setUp(self):
        self.transport = FakeTransport({
            "auth/currentUserInfo": {"transport": "ok", "code": 0, "result": {"userId": "1"}},
            "view/subjects": SUBJECTS,
            "view/booking/yyxh": SCHEDULE,
            "view/booking/yyxkzw": SEATS,
            "view/lesson/ckkb": {"transport": "ok", "code": 0, "result": []},
        })
        self.static = {
            "http://openlab.hitwh.edu.cn/dxwl/booking/": (
                '<script src="./_app.config.js?v=2.8.0-1774602871364">'
                "<script type=module crossorigin src=./assets/index.655b7571.js>"
            ),
            "http://openlab.hitwh.edu.cn/dxwl/booking/_app.config.js": '{"VITE_GLOB_API_URL":"/dxwl/StuApi"}',
        }

    def observation(self):
        return observe(self.transport, static_get=lambda url, timeout: self.static[url],
                       environment={"proxy": "127.0.0.1:7897"})

    def test_observation_feeds_parameters_from_earlier_reads(self):
        result = self.observation()
        self.assertEqual([], result.unavailable)
        self.assertEqual(
            [("view/booking/yyxh", {"id": 3001}),
             ("view/booking/yyxkzw", {"subjectId": 3001, "classDate": "2026-09-22", "timer": 2})],
            [call for call in self.transport.calls if call[0].startswith("view/booking/")],
        )

    def test_observation_records_shapes_versions_and_static_signals(self):
        snapshot_ = self.observation().snapshot
        self.assertEqual(("index.655b7571.js",), snapshot_.static_assets)
        self.assertEqual(("VITE_GLOB_API_URL",), snapshot_.static_config)
        self.assertEqual(
            ("data.lab.LabsAddress", "data.lab.LabsName", "data.labList[].TableNo",
             "data.labList[].TableShow", "data.labList[].id", "data.labList[].status", "msg"),
            snapshot_.endpoints["view/booking/yyxkzw"].response_fields,
        )
        self.assertEqual(("izPass", "subjectId", "subjectName"),
                         snapshot_.endpoints["view/subjects"].response_fields)
        self.assertEqual(("code", "message", "result", "timestamp"),
                         snapshot_.endpoints["view/subjects"].envelope)

    def test_observation_reads_the_shell_version_and_unquoted_assets(self):
        snapshot_ = self.observation().snapshot
        self.assertEqual("2.8.0-1774602871364", snapshot_.app_version)
        self.assertEqual(("index.655b7571.js",), snapshot_.static_assets)

    def test_business_failure_marks_the_endpoint_unavailable(self):
        transport = FakeTransport({
            "auth/currentUserInfo": {"transport": "ok", "code": 0, "result": {"userId": "1"}},
            "view/subjects": SUBJECTS,
            "view/lesson/ckkb": {"transport": "ok", "code": 0, "result": []},
            "view/booking/yyxh": {"transport": "ok", "code": 5000, "message": "您已经预约过此实验项目"},
        })
        result = observe(transport, static_get=lambda url, timeout: "")
        self.assertIn("view/booking/yyxh", result.unavailable)
        self.assertNotIn("view/booking/yyxh", result.snapshot.endpoints)
        self.assertIn("view/booking/yyxkzw", result.unavailable)

    def test_probe_prefers_an_experiment_that_is_not_booked_yet(self):
        transport = FakeTransport({
            "auth/currentUserInfo": {"transport": "ok", "code": 0, "result": {}},
            "view/subjects": {"transport": "ok", "code": 0, "result": [
                {"subjectId": 3001, "subjectName": "已约", "izPass": True},
                {"subjectId": 3002, "subjectName": "未约", "izPass": False}]},
            "view/lesson/ckkb": {"transport": "ok", "code": 0, "result": [{"subjectId": 3001}]},
            "view/booking/yyxh": SCHEDULE,
            "view/booking/yyxkzw": SEATS,
        })
        observe(transport, static_get=lambda url, timeout: "")
        forms = dict(transport.calls)
        self.assertEqual({"id": 3002}, forms["view/booking/yyxh"])
        self.assertEqual(3002, forms["view/booking/yyxkzw"]["subjectId"])

    def test_probe_tries_several_cells_before_giving_up_on_seat_shapes(self):
        calls = {"yyxkzw": 0}

        class FlakyTransport(FakeTransport):
            def call(self, path, form=None):
                if path == "view/booking/yyxkzw":
                    calls["yyxkzw"] += 1
                    if calls["yyxkzw"] == 1:
                        return {"transport": "ok", "code": 5000, "message": "没有可供选择的座位"}
                return super().call(path, form)

        transport = FlakyTransport({
            "auth/currentUserInfo": {"transport": "ok", "code": 0, "result": {}},
            "view/lesson/ckkb": {"transport": "ok", "code": 0, "result": []},
            "view/subjects": SUBJECTS,
            "view/booking/yyxh": {"transport": "ok", "code": 0, "result": {"data": [
                {"classDate": "2026-09-22", "eduWeek": 4, "dayWeek": "星期二",
                 "timerList": [{"timer": 2, "timerName": "第二大节", "startTime": "10:05:00"}]},
                {"classDate": "2026-09-23", "eduWeek": 4, "dayWeek": "星期三",
                 "timerList": [{"timer": 2, "timerName": "第二大节", "startTime": "10:05:00"}]}]}},
            "view/booking/yyxkzw": SEATS,
        })
        result = observe(transport, static_get=lambda url, timeout: "")
        self.assertEqual([], result.unavailable)
        self.assertIn("view/booking/yyxkzw", result.snapshot.endpoints)
        self.assertEqual(2, calls["yyxkzw"])

    def test_failed_static_fetch_degrades_to_unknown_without_aborting(self):
        result = observe(self.transport, static_get=lambda url, timeout: (_ for _ in ()).throw(OSError("dns")))
        self.assertEqual((), result.snapshot.static_assets)
        self.assertTrue(result.snapshot.endpoints)

    def test_unreadable_endpoints_are_reported_not_silently_dropped(self):
        transport = FakeTransport({"view/subjects": SUBJECTS})
        result = observe(transport, static_get=lambda url, timeout: "")
        self.assertIn("view/lesson/ckkb", result.unavailable)
        self.assertIn("view/subjects", result.snapshot.endpoints)
        self.assertNotIn("auth/currentUserInfo", result.snapshot.endpoints)

    def test_locked_fields_can_be_marked_without_losing_observations(self):
        snapshot_ = with_locked(self.observation().snapshot,
                                {"view/subjects": ["subjectId", "izPass"]})
        self.assertEqual(("izPass", "subjectId"), snapshot_.endpoints["view/subjects"].locked)
        self.assertEqual(("izPass", "subjectId", "subjectName"),
                         snapshot_.endpoints["view/subjects"].response_fields)


if __name__ == "__main__":
    unittest.main()
