import tempfile
from pathlib import Path

from course_selection.browser_observer import ObserverResult
from course_selection.gateway import UnconfirmedAcademicGateway
from course_selection.workbench import create_workbench_app


class FakeObserver:
    def checkpoint(self):
        events = [
            {"kind": "request", "method": "POST", "url_shape": "https://<host:a>/<path:2>",
             "target_identity": "target", "frame_identity": "frame", "loader_identity": "loader",
             "initiator_class": "frame", "resource_type": "xhr", "elapsed_ms": 10},
            {"kind": "response", "method": "POST", "url_shape": "https://<host:a>/<path:2>",
             "target_identity": "target", "resource_type": "xhr", "status": 200},
        ]
        return ObserverResult("complete", "checkpoint", {"trace_id": "trace-1", "events": events})

    def shutdown(self):
        return ObserverResult("disconnected", "shutdown")


def test_checkpoint_http_adapter_returns_ranked_diagnostic_candidates():
    with tempfile.TemporaryDirectory() as directory:
        app = create_workbench_app(
            Path(directory), gateway_factory=UnconfirmedAcademicGateway,
            browser_observer=FakeObserver(),
        )
        client = app.test_client()
        response = client.get("/api/browser-observer/checkpoint")
        assert response.status_code == 200
        payload = response.get_json()
        assert payload["data"]["candidates"][0]["method"] == "POST"
        assert payload["data"]["candidates"][0]["evidence_identity"]
        app.extensions["browser_observer"].shutdown()
        app.extensions["observation_service"].close()
        app.extensions["workspace_database"].close()
