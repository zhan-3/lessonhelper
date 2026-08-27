import unittest

from course_selection.manual_observation import ManualObservationPolicy, summarize_json_structure


class ManualObservationPolicyTests(unittest.TestCase):
    def test_official_authentication_post_is_allowed_but_lookalikes_are_blocked(self):
        policy = ManualObservationPolicy([])
        body = "username=student&password=redacted&execution=e1s1&submit=login"

        allowed, evidence = policy.inspect_request(
            "POST",
            "https://webvpn.hitwh.edu.cn/https/hash/authserver/login?service=academic",
            body,
            "application/x-www-form-urlencoded",
            "document",
        )
        lookalike, _ = policy.inspect_request(
            "POST",
            "https://evil.example/authserver/login?next=webvpn.hitwh.edu.cn",
            body,
            "application/x-www-form-urlencoded",
            "document",
        )

        self.assertTrue(allowed)
        self.assertFalse(evidence["blocked"])
        self.assertFalse(lookalike)
        self.assertNotIn("student", str(evidence))
        self.assertNotIn("redacted", str(evidence))

    def test_unknown_post_is_blocked_without_persisting_values(self):
        policy = ManualObservationPolicy([])

        allowed, evidence = policy.inspect_request(
            "POST",
            "https://webvpn.hitwh.edu.cn/api/resource/open?tab=academic",
            "student_number=20250001&action=query",
            "application/x-www-form-urlencoded",
            "xhr",
        )

        self.assertFalse(allowed)
        self.assertEqual(["action", "student_number"], evidence["field_names"])
        self.assertNotIn("20250001", str(evidence))
        self.assertTrue(evidence["blocked"])

    def test_only_exact_confirmed_post_signature_is_allowed(self):
        policy = ManualObservationPolicy([
            {
                "url": "https://webvpn.hitwh.edu.cn/api/resource/open",
                "field_names": ["action", "resource_id"],
                "query_field_names": [],
            }
        ])

        allowed, evidence = policy.inspect_request(
            "POST",
            "https://webvpn.hitwh.edu.cn/api/resource/open",
            '{"resource_id":"academic","action":"open"}',
            "application/json",
            "fetch",
        )
        changed_fields, _ = policy.inspect_request(
            "POST",
            "https://webvpn.hitwh.edu.cn/api/resource/open",
            '{"resource_id":"academic","action":"open","confirm":true}',
            "application/json",
            "fetch",
        )
        changed_query, _ = policy.inspect_request(
            "POST",
            "https://webvpn.hitwh.edu.cn/api/resource/open?retry=true",
            '{"resource_id":"academic","action":"open"}',
            "application/json",
            "fetch",
        )

        self.assertTrue(allowed)
        self.assertFalse(evidence["blocked"])
        self.assertFalse(changed_fields)
        self.assertFalse(changed_query)

    def test_response_summary_keeps_shape_without_values(self):
        summary = summarize_json_structure({
            "courses": [{"course_id": "secret-42", "name": "Example"}],
            "page_count": 3,
        })

        self.assertEqual("object", summary["type"])
        self.assertEqual("array", summary["fields"]["courses"]["type"])
        self.assertEqual("number", summary["fields"]["page_count"]["type"])
        self.assertNotIn("secret-42", str(summary))
        self.assertNotIn("Example", str(summary))

    def test_confirmed_post_can_restrict_a_field_to_exact_allowed_values(self):
        policy = ManualObservationPolicy([
            {
                "url": "https://example.test/xsxk/queryXsxkList",
                "query_field_names": [],
                "field_names": ["pageXklb", "pageXnxq", "token"],
                "allowed_values": {"pageXklb": ["allowed-code", "second-code"]},
                "integer_ranges": {"pageXnxq": [2020, 2030]},
            }
        ])

        allowed, evidence = policy.inspect_request(
            "POST",
            "https://example.test/xsxk/queryXsxkList",
            "pageXklb=allowed-code&pageXnxq=2026&token=secret",
            "application/x-www-form-urlencoded",
            "document",
        )
        denied, denied_evidence = policy.inspect_request(
            "POST",
            "https://example.test/xsxk/queryXsxkList",
            "pageXklb=denied-code&pageXnxq=2026&token=secret",
            "application/x-www-form-urlencoded",
            "document",
        )
        out_of_range, _ = policy.inspect_request(
            "POST",
            "https://example.test/xsxk/queryXsxkList",
            "pageXklb=allowed-code&pageXnxq=2031&token=secret",
            "application/x-www-form-urlencoded",
            "document",
        )

        self.assertTrue(allowed)
        self.assertFalse(denied)
        self.assertFalse(out_of_range)
        self.assertNotIn("allowed-code", str(evidence))
        self.assertNotIn("denied-code", str(denied_evidence))
        self.assertNotIn("secret", str(evidence))

    def test_get_navigation_is_observed_and_allowed(self):
        policy = ManualObservationPolicy([])

        allowed, evidence = policy.inspect_request(
            "GET", "https://webvpn.hitwh.edu.cn/#!/service", None, "", "document"
        )

        self.assertTrue(allowed)
        self.assertEqual("document", evidence["resource_type"])
        self.assertEqual([], evidence["field_names"])


if __name__ == "__main__":
    unittest.main()
