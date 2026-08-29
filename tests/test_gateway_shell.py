import unittest
from types import SimpleNamespace

from course_selection.gateway import PlaywrightAcademicGateway


class _Page:
    def __init__(self, url: str, *, fail_navigation: bool = False):
        self.url = url
        self.fail_navigation = fail_navigation
        self.closed = False
        self.front = False
        self.navigations: list[str] = []

    def is_closed(self):
        return self.closed

    def goto(self, url, **_kwargs):
        self.navigations.append(url)
        if self.fail_navigation:
            raise TimeoutError("page is unresponsive")
        self.url = url

    def bring_to_front(self):
        self.front = True

    def close(self):
        self.closed = True


class _Context:
    def __init__(self, pages):
        self.pages = pages
        self.created: list[_Page] = []

    def new_page(self):
        page = _Page("about:blank")
        self.pages.append(page)
        self.created.append(page)
        return page


class GatewayShellRecoveryTests(unittest.TestCase):
    def test_unresponsive_existing_shell_is_replaced_with_a_fresh_page(self):
        stale = _Page("http://127.0.0.1:5000/", fail_navigation=True)
        context = _Context([stale])
        gateway = PlaywrightAcademicGateway()
        gateway._session = SimpleNamespace(context=context)
        progress = []

        gateway.launch_shell(
            "http://127.0.0.1:5000",
            lambda state, details: progress.append((state, details)),
            lambda: False,
        )

        self.assertTrue(stale.closed)
        self.assertEqual(1, len(context.created))
        self.assertEqual(["http://127.0.0.1:5000/"], context.created[0].navigations)
        self.assertTrue(context.created[0].front)
        self.assertEqual("visible Chromium workbench opened", progress[-1][1]["message"])


if __name__ == "__main__":
    unittest.main()
