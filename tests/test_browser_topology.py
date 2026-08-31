import time

import pytest

from course_selection.browser_observer import BorrowedBrowserObserver
from course_selection.browser_observer_worker import BorrowedBrowserObserverWorker


class FakeFrame:
    def __init__(self, url, *, parent=None, page=None, name=""):
        self.url = url
        self.parent_frame = parent
        self.page = page
        self.name = name


class FakeWorker:
    def __init__(self, url):
        self.url = url
        self._impl_obj = type("Impl", (), {"_guid": f"worker:{url}"})()


class FakePage:
    def __init__(self, url, *, opener=None):
        self.url = url
        self._opener = opener
        self.frames = []
        self.workers = []

    def opener(self):
        return self._opener

    def on(self, *_args):
        return None

    def remove_listener(self, *_args):
        return None

    def wait_for_timeout(self, milliseconds):
        time.sleep(milliseconds / 1000)


class FakeContext:
    def __init__(self, pages):
        self.pages = pages
        for page in pages:
            page.context = self
        self.service_workers = [FakeWorker("https://worker.example/service.js")]
        self._impl_obj = type("Impl", (), {"_guid": "fixture-context"})()

    def on(self, *_args):
        return None

    def remove_listener(self, *_args):
        return None


class FakeBrowser:
    def __init__(self, pages):
        self.contexts = [FakeContext(pages)]

    def is_connected(self):
        return True


def inventory_for(pages):
    observer = BorrowedBrowserObserver(session_id="topology-test")
    observer._browser = FakeBrowser(pages)
    return observer.inventory().data["targets"]


def test_inventory_preserves_frame_and_popup_relationships_without_order_dependence():
    parent = FakePage("https://portal.example/home")
    main = FakeFrame(parent.url, page=parent)
    child = FakeFrame("https://frame.example/app", parent=main, page=parent)
    nested = FakeFrame("https://other.example/detail", parent=child, page=parent)
    sibling = FakeFrame("https://frame.example/app", parent=main, page=parent)
    parent.frames = [main, child, nested, sibling]
    parent.workers = [FakeWorker("https://worker.example/page.js")]
    popup = FakePage("about:blank", opener=parent)
    popup.frames = [FakeFrame("about:blank", page=popup)]

    first = inventory_for([popup, parent])
    second = inventory_for([parent, popup])
    assert first == second
    context_target = next(item for item in first if item["kind"] == "context")
    top_level = next(item for item in first if item["kind"] == "page" and item["relationship"] == "top_level")
    assert top_level["parent_identity"] == context_target["target_identity"]
    popup_target = next(item for item in first if item["kind"] == "page" and item["relationship"] == "popup")
    assert popup_target["navigation_state"] == "unresolved_blank"
    assert popup_target["parent_identity"]
    frame_targets = [item for item in first if item["kind"] == "frame"]
    assert len(frame_targets) == 3
    assert len({item["target_identity"] for item in frame_targets}) == 3
    assert all(item["parent_identity"] for item in frame_targets)
    assert all(item["capabilities"] == ["document", "network"] for item in frame_targets)
    worker_targets = [item for item in first if item["kind"] in {"worker", "service_worker"}]
    assert len(worker_targets) == 2
    assert all(item["parent_identity"] for item in worker_targets)

    observer = BorrowedBrowserObserver(session_id="request-provenance")
    request = type("Request", (), {"frame": main, "resource_type": "document"})()
    context = observer._request_context(request)
    page_target = observer._page_target(parent)
    assert context["target_identity"] == page_target["target_identity"]
    assert context["frame_identity"] != context["target_identity"]


@pytest.mark.parametrize(
    ("depth", "origin", "frame_name", "reverse"),
    [(1, "https://one.example/app", "business", False), (3, "https://cross.example/app", "renamed", True)],
)
def test_topology_variants_do_not_depend_on_names_depth_origin_or_order(depth, origin, frame_name, reverse):
    page = FakePage("https://portal.example/home")
    main = FakeFrame(page.url, page=page)
    frames = [main]
    parent = main
    for index in range(depth):
        child = FakeFrame(f"{origin}/{index}", parent=parent, page=page, name=f"{frame_name}-{index}")
        frames.append(child)
        parent = child
    page.frames = frames
    popup = FakePage("about:blank", opener=page)
    popup.frames = [FakeFrame("about:blank", page=popup)]
    pages = [popup, page] if reverse else [page, popup]
    first = inventory_for(pages)
    unresolved = next(item for item in first if item["relationship"] == "popup")
    popup.url = "https://application.example/ready"
    popup.frames[0].url = popup.url
    second = inventory_for(pages)
    committed = next(item for item in second if item["relationship"] == "popup")
    assert unresolved["target_identity"] == committed["target_identity"]
    assert unresolved["navigation_state"] == "unresolved_blank"
    assert committed["navigation_state"] == "committed"
    assert len([item for item in second if item["kind"] == "frame"]) == depth


def test_runtime_budget_ends_trace_with_typed_partial_outcome():
    page = FakePage("https://portal.example/home")
    page.frames = [FakeFrame(page.url, page=page)]
    core = BorrowedBrowserObserver(session_id="runtime-budget", max_runtime_seconds=0.02)
    core._browser = FakeBrowser([page])
    observer = BorrowedBrowserObserverWorker(core)
    try:
        assert observer.start_observation().status == "observing"
        time.sleep(0.08)
        result = observer.checkpoint()
        assert result.status == "partial"
        assert "runtime_budget_exhausted" in result.data["missing_evidence"]
        assert result.next_actions
    finally:
        observer.shutdown()
