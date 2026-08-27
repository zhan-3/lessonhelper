"""Persist JSON network exchanges and rank graduation-related API candidates."""

from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .sanitizer import sanitize_data, sanitize_request_body, sanitize_url


@dataclass(frozen=True)
class Candidate:
    capture_id: str
    score: int
    reasons: tuple[str, ...]
    method: str
    status: int
    url: str
    sanitized_response: str


_SIGNALS: tuple[tuple[str, int, str], ...] = (
    ("培养方案", 10, "contains 培养方案"),
    ("方案完成", 10, "contains 方案完成"),
    ("毕业", 8, "contains 毕业"),
    ("学业完成", 8, "contains 学业完成"),
    ("已修", 7, "contains 已修"),
    ("学分", 5, "contains 学分"),
    ("课程类别", 5, "contains 课程类别"),
    ("课程性质", 5, "contains 课程性质"),
    ("training", 7, "contains training"),
    ("graduate", 7, "contains graduate"),
    ("graduation", 7, "contains graduation"),
    ("program", 5, "contains program"),
    ("scheme", 5, "contains scheme"),
    ("credit", 4, "contains credit"),
    ("course", 3, "contains course"),
    ("curriculum", 7, "contains curriculum"),
)


def score_candidate(url: str, response_data: Any) -> tuple[int, tuple[str, ...]]:
    searchable = f"{url}\n{json.dumps(response_data, ensure_ascii=False)}".lower()
    score = 0
    reasons: list[str] = []
    for signal, weight, reason in _SIGNALS:
        if signal.lower() in searchable:
            score += weight
            reasons.append(reason)

    if isinstance(response_data, (dict, list)):
        score += 1
        reasons.append("structured JSON")

    return score, tuple(reasons)


class CaptureStore:
    """Store raw response bodies privately and write sanitized searchable indexes."""

    def __init__(self, root: Path):
        self.root = root
        self.raw_dir = root / "raw"
        self.sanitized_dir = root / "sanitized"
        self.raw_dir.mkdir(parents=True, exist_ok=True)
        self.sanitized_dir.mkdir(parents=True, exist_ok=True)
        self._counter = self._next_counter()
        self._candidates: list[Candidate] = []

    def _next_counter(self) -> int:
        existing = [
            int(path.stem)
            for path in self.raw_dir.glob("*.json")
            if re.fullmatch(r"\d+", path.stem)
        ]
        return max(existing, default=0) + 1

    def save_json_exchange(
        self,
        *,
        url: str,
        method: str,
        status: int,
        content_type: str,
        request_body: str | None,
        response_data: Any,
        captured_at: str | None = None,
    ) -> Candidate:
        capture_id = f"{self._counter:06d}"
        self._counter += 1

        raw_path = self.raw_dir / f"{capture_id}.json"
        sanitized_path = self.sanitized_dir / f"{capture_id}.json"

        raw_path.write_text(
            json.dumps(response_data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        sanitized = sanitize_data(response_data)
        sanitized_path.write_text(
            json.dumps(sanitized, ensure_ascii=False, indent=2), encoding="utf-8"
        )

        safe_url = sanitize_url(url)
        score, reasons = score_candidate(safe_url, sanitized)
        candidate = Candidate(
            capture_id=capture_id,
            score=score,
            reasons=reasons,
            method=method,
            status=status,
            url=safe_url,
            sanitized_response=str(sanitized_path.relative_to(self.root)),
        )
        self._candidates.append(candidate)

        index_entry = {
            "capture_id": capture_id,
            "captured_at": captured_at
            or datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "method": method,
            "status": status,
            "content_type": content_type,
            "url": safe_url,
            "request_body": sanitize_request_body(request_body) if request_body else None,
            "raw_response": str(raw_path.relative_to(self.root)),
            "sanitized_response": str(sanitized_path.relative_to(self.root)),
            "candidate_score": score,
            "candidate_reasons": reasons,
        }
        with (self.root / "index.jsonl").open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(index_entry, ensure_ascii=False) + "\n")

        return candidate

    def write_candidates(self) -> Path:
        path = self.root / "candidates.json"
        ranked = sorted(
            self._candidates,
            key=lambda candidate: (-candidate.score, candidate.capture_id),
        )
        path.write_text(
            json.dumps([asdict(item) for item in ranked], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return path
