import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from course_selection.discovery import DiscoveryReport
from course_selection.gateway import PlaywrightAcademicGateway


class _Context:
    def on(self, *_args):
        return None

    def route(self, *_args):
        return None


class SelectionSnapshotTests(unittest.TestCase):
    def test_gateway_consumes_in_memory_payload_without_reading_diagnostic_json(self):
        payload = {
            "queries": [
                {
                    "category": "szhx",
                    "complete": True,
                    "sections": [{"identity": "section-1", "name": "Test"}],
                }
            ]
        }
        report = DiscoveryReport(
            target="selection",
            target_found=True,
            clicks=0,
            captures=0,
            blocked_requests=0,
            candidates_path=Path("candidates.json"),
            click_log_path=Path("clicks.json"),
            # Deliberately point at a file that does not exist.  The refresh
            # result must still be produced from the structured payload.
            selection_query_path=Path("missing-selection-query.json"),
            selection_query_payload=payload,
        )

        class Navigator:
            def __init__(self, **_kwargs):
                pass

            def _handle_response(self, *_args):
                return None

            def _guard_route(self, *_args):
                return None

            def run(self, *_args, **_kwargs):
                return report

        with tempfile.TemporaryDirectory() as directory:
            gateway = PlaywrightAcademicGateway(workspace_root=directory)
            gateway._session = SimpleNamespace(context=_Context())
            with patch("course_selection.discovery.InterfaceDiscovery", Navigator):
                result = gateway.refresh_selection(
                    {
                        "allowed_categories": ("szhx",),
                        "allowed_windows": {},
                        "semester_label": "2026-1",
                    },
                    lambda *_args: None,
                    lambda: False,
                )

        self.assertEqual("complete", result["status"])
        self.assertEqual("section-1", result["sections"][0]["identity"])


if __name__ == "__main__":
    unittest.main()
