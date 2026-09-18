import unittest

from course_selection.lab_transport import (
    CONTENT_TYPE,
    TOKEN_HEADER,
    BrowserLabTransport,
    HttpLabTransport,
    RoutedLabTransport,
    install_token_hook,
    page_token,
)


class Recorder:
    """Stands in for urllib so tests can assert the exact wire request."""

    def __init__(self, *, status=200, text='{"code":0,"message":"请求成功","result":[]}', error=None):
        self.calls = []
        self.status, self.text, self.error = status, text, error

    def __call__(self, url, data, headers, timeout):
        self.calls.append({"url": url, "data": data, "headers": dict(headers), "timeout": timeout})
        if self.error is not None:
            raise self.error
        return self.status, self.text


class HttpTransportTests(unittest.TestCase):
    def test_sends_form_encoded_post_with_token_header_and_no_cookie(self):
        recorder = Recorder()
        transport = HttpLabTransport("dxwl", "token-value", fetch=recorder)

        payload = transport.call("view/booking/yyxkzw", {"subjectId": 3001, "timer": 2})

        self.assertEqual("ok", payload["transport"])
        self.assertEqual(0, payload["code"])
        call = recorder.calls[0]
        self.assertEqual("http://openlab.hitwh.edu.cn/dxwl/StuApi/view/booking/yyxkzw", call["url"])
        self.assertEqual(b"subjectId=3001&timer=2", call["data"])
        self.assertEqual("token-value", call["headers"][TOKEN_HEADER])
        self.assertEqual(CONTENT_TYPE, call["headers"]["Content-Type"])
        self.assertNotIn("Cookie", call["headers"])

    def test_transport_failure_is_typed_instead_of_raised(self):
        transport = HttpLabTransport("dxwl", "t", fetch=Recorder(error=OSError("no route")))
        payload = transport.call("view/subjects")
        self.assertEqual("error", payload["transport"])
        self.assertIn("OSError", payload["detail"])

    def test_non_json_response_is_reported_as_such(self):
        transport = HttpLabTransport("dxwl", "t", fetch=Recorder(text="<html>waf</html>"))
        payload = transport.call("view/subjects")
        self.assertEqual("non_json", payload["transport"])

    def test_dropped_parameters_are_not_sent(self):
        recorder = Recorder()
        HttpLabTransport("dxwl", "t", fetch=recorder).call("view/booking/yyxkzw", {"timer": 2, "classDate": None})
        self.assertEqual(b"timer=2", recorder.calls[0]["data"])


class BrowserTransportTests(unittest.TestCase):
    def test_browser_backend_posts_inside_the_page(self):
        recorded = []

        class Page:
            def evaluate(self, script, argument=None):
                if argument is None:
                    return "installed"
                recorded.append(argument)
                return {"transport": "ok", "code": 0, "message": "请求成功", "result": True}

        payload = BrowserLabTransport(Page(), "dxwl", "tok", pause=0).call("view/lesson/ckkb", {})
        self.assertTrue(payload["result"])
        path, form, token = recorded[-1]
        self.assertEqual("/dxwl/StuApi/view/lesson/ckkb", path)
        self.assertEqual({}, form)
        self.assertEqual("tok", token)

    def test_hook_install_and_token_read_use_the_page(self):
        class Page:
            def __init__(self):
                self.scripts = []

            def evaluate(self, script, argument=None):
                self.scripts.append(script)
                return "installed" if len(self.scripts) == 1 else "captured-token"

        page = Page()
        self.assertEqual("installed", install_token_hook(page))
        self.assertEqual("captured-token", page_token(page))
        self.assertIn("vctchauthorization", page.scripts[0])


class RoutedTransportTests(unittest.TestCase):
    def test_falls_back_only_when_the_network_path_fails(self):
        primary = HttpLabTransport("dxwl", "t", fetch=Recorder(error=OSError("dns")))
        fallback = HttpLabTransport("dxwl", "t", fetch=Recorder(text='{"code":0,"result":[1]}'))
        routed = RoutedLabTransport(primary, fallback)

        payload = routed.call("view/subjects")

        self.assertEqual("fallback", routed.last_used)
        self.assertEqual([1], payload["result"])

    def test_business_failure_is_never_replayed_on_the_other_backend(self):
        primary = Recorder(text='{"code":5000,"message":"没有可供选择的座位"}')
        fallback = Recorder(text='{"code":0,"result":[1]}')
        routed = RoutedLabTransport(HttpLabTransport("dxwl", "t", fetch=primary),
                                    HttpLabTransport("dxwl", "t", fetch=fallback))

        payload = routed.call("view/booking/yyxkzw", {"subjectId": 1})

        self.assertEqual("primary", routed.last_used)
        self.assertEqual(5000, payload["code"])
        self.assertEqual([], fallback.calls)


if __name__ == "__main__":
    unittest.main()
