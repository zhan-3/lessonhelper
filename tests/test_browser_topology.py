from course_selection.browser_observer import BorrowedBrowserObserver


class FakeFrame:
    def __init__(self, url, *, parent=None, page=None):
        self.url = url
        self.parent_frame = parent
        self.page = page


class FakePage:
    def __init__(self, url, *, opener=None):
        self.url = url
        self._opener = opener
        self.frames = []
        self.workers = []

    def opener(self):
        return self._opener


class FakeContext:
    def __init__(self, pages):
        self.pages = pages
        self.service_workers = []


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
    popup = FakePage("about:blank", opener=parent)
    popup.frames = [FakeFrame("about:blank", page=popup)]

    first = inventory_for([popup, parent])
    second = inventory_for([parent, popup])
    assert first == second
    popup_target = next(item for item in first if item["kind"] == "page" and item["relationship"] == "popup")
    assert popup_target["navigation_state"] == "unresolved_blank"
    assert popup_target["parent_identity"]
    frame_targets = [item for item in first if item["kind"] == "frame"]
    assert len(frame_targets) == 3
    assert len({item["target_identity"] for item in frame_targets}) == 3
    assert all(item["parent_identity"] for item in frame_targets)
    assert all(item["capabilities"] == ["document", "network"] for item in frame_targets)

    observer = BorrowedBrowserObserver(session_id="request-provenance")
    request = type("Request", (), {"frame": main, "resource_type": "document"})()
    context = observer._request_context(request)
    page_target = observer._page_target(parent)
    assert context["target_identity"] == page_target["target_identity"]
    assert context["frame_identity"] != context["target_identity"]
