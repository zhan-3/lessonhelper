"""Contract snapshots for openlab center apps: record, fingerprint, diff.

A snapshot is the small set of facts our code depends on: which endpoints exist,
what we send, what comes back, which business codes appear, and which static
assets and environment the observation belongs to.  Baselines live in the
repository so a school-side change shows up as a reviewable diff; observations
belong to one machine and stay local.

The comparison is deliberately narrow - version signal, endpoint set, and
response top-level field names - because that is where real changes show up
while type- and value-level churn would only produce noise.
"""

from __future__ import annotations

import hashlib
import json
import re
import urllib.request
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from typing import Any

from .lab_transport import DEFAULT_ORIGIN, TOKEN_HEADER, LabTransport

SCHEMA_VERSION = 1

# Only read endpoints are observed.  `view/booking/doyyxkzw` is a write and is
# never part of a contract snapshot.
ALLOWLIST: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("auth/currentUserInfo", ()),
    ("view/lesson/ckkb", ()),
    ("view/subjects", ()),
    ("view/booking/yyxh", ("id",)),
    ("view/booking/yyxkzw", ("subjectId", "classDate", "timer")),
)

REBASELINE = "rebaseline"
UNAVAILABLE = "unavailable"
BREAKING = "breaking"
ADDITIVE = "additive"
NOTICE = "notice"
COSMETIC = "cosmetic"

BLOCKING = frozenset({BREAKING, UNAVAILABLE})

# The booking shell writes unquoted attributes (src=./assets/index.<hash>.js),
# so both quoted and bare values have to be accepted.
_ASSET_PATTERN = re.compile(
    r"""(?:src|href)=(?:"[^"]*?"|'[^']*?'|[^\s>]+)""",
)
_ASSET_FILE_PATTERN = re.compile(r"/assets/([A-Za-z0-9._-]+)")
_VERSION_PATTERN = re.compile(r"""_app\.config\.js\?v=([A-Za-z0-9._-]+)""")
_CONFIG_KEY_PATTERN = re.compile(r"""["'](VITE_[A-Z0-9_]+)["']\s*:""")


@dataclass(frozen=True)
class EndpointContract:
    endpoint: str
    request_fields: tuple[str, ...] = ()
    response_fields: tuple[str, ...] = ()
    locked: tuple[str, ...] = ()
    envelope: tuple[str, ...] = ()
    codes: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return {
            "request_fields": list(self.request_fields),
            "response_fields": list(self.response_fields),
            "locked": list(self.locked),
            "envelope": list(self.envelope),
            "codes": list(self.codes),
        }

    @classmethod
    def from_dict(cls, endpoint: str, payload: Mapping[str, Any]) -> EndpointContract:
        return cls(
            endpoint=endpoint,
            request_fields=tuple(payload.get("request_fields") or ()),
            response_fields=tuple(payload.get("response_fields") or ()),
            locked=tuple(payload.get("locked") or ()),
            envelope=tuple(payload.get("envelope") or ()),
            codes=tuple(str(value) for value in payload.get("codes") or ()),
        )


@dataclass(frozen=True)
class ContractSnapshot:
    center: str
    channel: str = "direct"
    origin: str = DEFAULT_ORIGIN
    app_version: str = ""
    header_required: tuple[str, ...] = (TOKEN_HEADER,)
    static_assets: tuple[str, ...] = ()
    static_config: tuple[str, ...] = ()
    environment: Mapping[str, str] = field(default_factory=dict)
    endpoints: Mapping[str, EndpointContract] = field(default_factory=dict)
    unavailable: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema_version": SCHEMA_VERSION,
            "center": self.center,
            "channel": self.channel,
            "origin": self.origin,
            "app_version": self.app_version,
            "header_required": list(self.header_required),
            "static_assets": list(self.static_assets),
            "static_config": list(self.static_config),
            "environment": dict(sorted(self.environment.items())),
            "endpoints": {name: item.to_dict() for name, item in sorted(self.endpoints.items())},
            "unavailable": list(self.unavailable),
        }

    def to_json(self) -> str:
        return json.dumps(self.to_dict(), ensure_ascii=False, indent=2) + "\n"

    @classmethod
    def from_dict(cls, payload: Mapping[str, Any]) -> ContractSnapshot:
        version = int(payload.get("schema_version") or 0)
        if version != SCHEMA_VERSION:
            raise ValueError(f"unsupported contract schema_version: {version}")
        return cls(
            center=str(payload["center"]),
            channel=str(payload.get("channel") or "direct"),
            origin=str(payload.get("origin") or DEFAULT_ORIGIN),
            app_version=str(payload.get("app_version") or ""),
            header_required=tuple(payload.get("header_required") or (TOKEN_HEADER,)),
            static_assets=tuple(payload.get("static_assets") or ()),
            static_config=tuple(payload.get("static_config") or ()),
            environment={
                str(k): str(v)
                for k, v in (payload.get("environment") or {}).items()
            },
            endpoints={
                str(name): EndpointContract.from_dict(str(name), body)
                for name, body in (payload.get("endpoints") or {}).items()
            },
            unavailable=tuple(payload.get("unavailable") or ()),
        )

    @classmethod
    def from_json(cls, text: str) -> ContractSnapshot:
        return cls.from_dict(json.loads(text))


def fingerprint(snapshot: ContractSnapshot) -> str:
    """Stable hash of the parts our parsing depends on, ignoring environment noise."""
    canonical = {
        "endpoints": {
            name: {"response_fields": sorted(item.response_fields), "locked": sorted(item.locked)}
            for name, item in snapshot.endpoints.items()
        },
        "origin": snapshot.origin,
    }
    digest = hashlib.sha256(json.dumps(canonical, sort_keys=True).encode("utf-8"))
    return digest.hexdigest()[:16]


@dataclass(frozen=True)
class DriftItem:
    severity: str
    subject: str
    detail: str


@dataclass
class DriftReport:
    baseline_center: str
    current_center: str
    items: list[DriftItem] = field(default_factory=list)

    @property
    def blocks(self) -> bool:
        return any(item.severity in BLOCKING for item in self.items)

    def count(self, severity: str) -> int:
        return sum(1 for item in self.items if item.severity == severity)

    def to_dict(self) -> dict[str, Any]:
        return {
            "center": self.current_center,
            "blocks": self.blocks,
            "items": [
                {"severity": item.severity, "subject": item.subject, "detail": item.detail}
                for item in self.items
            ],
        }

    def describe(self) -> list[str]:
        order = {REBASELINE: 0, UNAVAILABLE: 1, BREAKING: 2, ADDITIVE: 3, NOTICE: 4, COSMETIC: 5}
        lines = [
            f"  [{item.severity:11}] {item.subject}: {item.detail}"
            for item in sorted(self.items, key=lambda item: (order.get(item.severity, 9), item.subject))
        ]
        return lines or ["  无变化"]


def diff(baseline: ContractSnapshot, current: ContractSnapshot) -> DriftReport:
    """Compare two snapshots, grading each difference by what it costs us."""
    report = DriftReport(baseline_center=baseline.center, current_center=current.center)

    if baseline.channel != current.channel:
        report.items.append(DriftItem(REBASELINE, "channel",
                                      f"{baseline.channel} → {current.channel}：应切换基线而不是报漂移"))
    if baseline.origin != current.origin:
        report.items.append(DriftItem(REBASELINE, "origin",
                                      f"{baseline.origin} → {current.origin}：凭据按 origin 绑定"))
    for key in sorted(set(baseline.environment) & set(current.environment)):
        before, after = baseline.environment[key], current.environment[key]
        if before != after:
            severity = NOTICE if key in {"proxy", "dns_source"} else ADDITIVE
            report.items.append(DriftItem(severity, f"environment.{key}", f"{before or '未记录'} → {after or '未记录'}"))

    for name, after in current.endpoints.items():
        before = baseline.endpoints.get(name)
        if before is None:
            report.items.append(DriftItem(ADDITIVE, name, "新增端点"))
            continue
        _diff_endpoint(report, before, after)

    readable = set(current.endpoints)
    for name in sorted(set(baseline.endpoints) - readable):
        if name in set(current.unavailable):
            report.items.append(DriftItem(UNAVAILABLE, name, "本次未能读到该端点"))
        else:
            report.items.append(DriftItem(BREAKING, name, "端点消失"))

    if baseline.app_version != current.app_version:
        report.items.append(DriftItem(COSMETIC, "app_version",
                                      f"{baseline.app_version or '未记录'} → {current.app_version or '未记录'}"))
    if tuple(baseline.static_assets) != tuple(current.static_assets):
        report.items.append(DriftItem(COSMETIC, "static_assets", "前端构建产物变化，建议做一次完整检查"))

    before_config, after_config = set(baseline.static_config), set(current.static_config)
    for key in sorted(after_config - before_config):
        report.items.append(DriftItem(ADDITIVE, f"config.{key}", "静态配置新增键"))
    for key in sorted(before_config - after_config):
        report.items.append(DriftItem(BREAKING, f"config.{key}", "静态配置键消失"))

    return report


def _diff_endpoint(report: DriftReport, before: EndpointContract, after: EndpointContract) -> None:
    if not after.response_fields:
        # An empty result (for example a student with no bookings yet) cannot
        # prove the element shape.  That only matters if the baseline did have a
        # claim to verify; otherwise there is nothing to compare.
        if before.response_fields:
            report.items.append(DriftItem(UNAVAILABLE, after.endpoint, "本次未能读到响应字段"))
        return
    locked = set(before.locked)
    before_fields, after_fields = set(before.response_fields), set(after.response_fields)
    for name in sorted(after_fields - before_fields):
        report.items.append(DriftItem(ADDITIVE, f"{after.endpoint}.{name}", "新增响应字段"))
    for name in sorted(before_fields - after_fields):
        severity = BREAKING if name in locked else ADDITIVE
        report.items.append(DriftItem(severity, f"{after.endpoint}.{name}",
                                      "必需字段消失" if name in locked else "响应字段消失"))
    for name in sorted(set(before.request_fields) - set(after.request_fields)):
        report.items.append(DriftItem(BREAKING, f"{after.endpoint}.{name}", "请求字段消失"))
    new_codes = sorted(set(after.codes) - set(before.codes))
    for code in new_codes:
        report.items.append(DriftItem(ADDITIVE, f"{after.endpoint}.code[{code}]", "新出现的业务码"))


def _field_paths(value: Any, prefix: str = "", depth: int = 3) -> tuple[str, ...]:
    """Dotted paths of the fields we read, so a change inside a nested list is visible."""
    if depth < 0:
        return ()
    if isinstance(value, Mapping):
        found: list[str] = []
        for key in sorted(str(name) for name in value):
            child = value[key]
            path = f"{prefix}{key}"
            if isinstance(child, Mapping):
                found.extend(_field_paths(child, f"{path}.", depth - 1))
            elif isinstance(child, Sequence) and not isinstance(child, (str, bytes)):
                for item in child[:1]:
                    if isinstance(item, Mapping):
                        found.extend(_field_paths(item, f"{path}[].", depth - 1))
                    else:
                        found.append(path)
                if not list(child[:1]):
                    found.append(path)
            else:
                found.append(path)
        return tuple(sorted(set(found)))
    if isinstance(value, Sequence) and not isinstance(value, (str, bytes)):
        found = []
        for item in value[:1]:
            if isinstance(item, Mapping):
                found.extend(_field_paths(item, prefix, depth - 1))
        return tuple(sorted(set(found)))
    return ()


def _asset_names(html: str) -> tuple[str, ...]:
    names: set[str] = set()
    for attribute in _ASSET_PATTERN.findall(html):
        names.update(_ASSET_FILE_PATTERN.findall(attribute))
    return tuple(sorted(names))


def _app_version(html: str) -> str:
    found = _VERSION_PATTERN.search(html)
    return found.group(1) if found else ""


def _config_keys(text: str) -> tuple[str, ...]:
    return tuple(sorted(set(_CONFIG_KEY_PATTERN.findall(text))))


def _default_get(url: str, timeout: float) -> str:
    with urllib.request.urlopen(url, timeout=timeout) as response:  # fixed campus origin
        return response.read().decode("utf-8", "replace")


def _optional_get(fetch_text: Callable[[str, float], str], url: str, timeout: float) -> str:
    """Static signals are best effort: a failure must not abort the observation."""
    try:
        return fetch_text(url, timeout)
    except Exception:
        return ""


def _forms_for(
    endpoint: str, request_fields: Sequence[str], context: Mapping[str, Any]
) -> list[dict[str, Any]]:
    """Parameter sets to try for one endpoint, most promising first."""
    if endpoint == "view/booking/yyxkzw":
        cells = context.get("cells") or []
        return [
            {"subjectId": cell[0], "classDate": cell[1], "timer": cell[2]}
            for cell in cells
        ]
    form = {name: context.get(name) for name in request_fields}
    if any(value is None for value in form.values()):
        return []
    return [form]


def _candidate_cells(subject_id: Any, schedule: Mapping[str, Any], limit: int = 3) -> list[tuple[Any, Any, Any]]:
    """A few (subject, date, timer) cells to try; a single full slot proves nothing."""
    cells: list[tuple[Any, Any, Any]] = []
    for row in (schedule.get("data") or [])[:limit]:
        for timer in (row.get("timerList") or [])[:1]:
            cells.append((subject_id, row.get("classDate"), timer.get("timer")))
    return cells[:limit]


@dataclass
class Observation:
    snapshot: ContractSnapshot
    unavailable: list[str] = field(default_factory=list)


def observe(
    transport: LabTransport,
    *,
    origin: str = DEFAULT_ORIGIN,
    channel: str = "direct",
    app_version: str = "",
    environment: Mapping[str, str] | None = None,
    static_get: Callable[[str, float], str] | None = None,
    timeout: float = 15.0,
) -> Observation:
    """Read the allowlisted endpoints once and describe what came back."""
    fetch_text = static_get or _default_get
    center = transport.center
    static_assets: tuple[str, ...] = ()
    static_config: tuple[str, ...] = ()
    shell = _optional_get(fetch_text, f"{origin}/{center}/booking/", timeout)
    static_assets = _asset_names(shell)
    static_config = _config_keys(
        _optional_get(fetch_text, f"{origin}/{center}/booking/_app.config.js", timeout)
    )
    version = app_version or _app_version(shell)

    endpoints: dict[str, EndpointContract] = {}
    unavailable: list[str] = []
    context: dict[str, Any] = {}

    for endpoint, request_fields in ALLOWLIST:
        forms = _forms_for(endpoint, request_fields, context)
        if not forms:
            unavailable.append(endpoint)
            continue
        payload, result = {}, None
        for form in forms:
            payload = transport.call(endpoint, form)
            if payload.get("transport") == "ok" and payload.get("code") == 0:
                result = payload.get("result")
                break
        if payload.get("transport") != "ok" or payload.get("code") != 0:
            unavailable.append(endpoint)
            continue
        envelope = tuple(sorted(str(k) for k in payload if k not in {"transport", "status"}))
        endpoints[endpoint] = EndpointContract(
            endpoint=endpoint,
            request_fields=request_fields,
            response_fields=_field_paths(result),
            envelope=envelope,
            codes=(str(payload.get("code")),) if payload.get("code") is not None else (),
        )
        if endpoint == "view/lesson/ckkb" and isinstance(result, Sequence):
            context["booked"] = {row.get("subjectId") for row in result if isinstance(row, Mapping)}
        if endpoint == "view/subjects" and isinstance(result, Sequence) and result:
            booked = context.get("booked") or set()
            free = [row for row in result if isinstance(row, Mapping) and row.get("subjectId") not in booked]
            context["id"] = ((free or list(result))[0] or {}).get("subjectId")
        if endpoint == "view/booking/yyxh" and isinstance(result, Mapping):
            context["cells"] = _candidate_cells(context.get("id"), result)

    snapshot = ContractSnapshot(
        center=center,
        channel=channel,
        origin=origin,
        app_version=version,
        static_assets=static_assets,
        static_config=static_config,
        environment=dict(environment or {}),
        endpoints=endpoints,
        unavailable=tuple(sorted(unavailable)),
    )
    return Observation(snapshot=snapshot, unavailable=unavailable)


def without_environment(snapshot: ContractSnapshot) -> ContractSnapshot:
    """Baseline copy: the shared contract carries no machine-local environment."""
    return ContractSnapshot(
        center=snapshot.center,
        channel=snapshot.channel,
        origin=snapshot.origin,
        app_version=snapshot.app_version,
        header_required=snapshot.header_required,
        static_assets=snapshot.static_assets,
        static_config=snapshot.static_config,
        environment={},
        endpoints=snapshot.endpoints,
        unavailable=snapshot.unavailable,
    )


def with_locked(snapshot: ContractSnapshot, locked: Mapping[str, Iterable[str]]) -> ContractSnapshot:
    """Return a copy with the human-chosen required fields marked."""
    endpoints = {
        name: EndpointContract(
            endpoint=item.endpoint,
            request_fields=item.request_fields,
            response_fields=item.response_fields,
            locked=tuple(sorted(locked.get(name, item.locked))),
            envelope=item.envelope,
            codes=item.codes,
        )
        for name, item in snapshot.endpoints.items()
    }
    return ContractSnapshot(
        center=snapshot.center,
        channel=snapshot.channel,
        origin=snapshot.origin,
        app_version=snapshot.app_version,
        header_required=snapshot.header_required,
        static_assets=snapshot.static_assets,
        static_config=snapshot.static_config,
        environment=snapshot.environment,
        endpoints=endpoints,
        unavailable=snapshot.unavailable,
    )
