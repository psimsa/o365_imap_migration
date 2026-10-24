"""Tier 1 performance tests: keep-alive sessions, inline attachments, $batch
fallback for large files, $select/$expand request shape, and 504 page shrink."""

import unittest
from unittest import mock

from tests import support

mig = support.mig


class RequestShapeTests(unittest.TestCase):
    def setUp(self):
        support.install_fakes()
        self.cfg = support.make_cfg(migrate_attachments=True)

    def test_message_pages_use_select_and_expand(self):
        list(mig.iter_source_emails(self.cfg, "u@example.com", "token"))
        page_calls = [(url, params) for url, params in support.FakeRequests.calls
                      if url.endswith("/messages") and "/next/" not in url]
        self.assertTrue(page_calls)
        for url, params in page_calls:
            self.assertIn("$select", params, f"missing $select in {url} {params}")
            self.assertIn("$expand", params, f"missing $expand in {url}")
            self.assertEqual(params["$expand"], mig.MESSAGE_ATTACHMENT_EXPAND)
            for prop in ("id", "subject", "body", "receivedDateTime", "attachments"):
                self.assertIn(prop, params["$select"])
        # the nextLink requests carry no params (query is in the URL)
        next_calls = [(url, params) for url, params in support.FakeRequests.calls
                      if "/next/" in url]
        self.assertTrue(next_calls)
        self.assertTrue(all(p == {} for _, p in next_calls),
                        f"nextLink calls must not repeat params: {next_calls}")

    def test_no_per_message_attachment_requests(self):
        messages = list(mig.iter_source_emails(self.cfg, "u@example.com", "token"))
        self.assertEqual(len(messages), 12)
        attachment_list_calls = [
            url for url, _ in support.FakeRequests.calls
            if url.endswith("/attachments")]
        self.assertEqual(attachment_list_calls, [],
                         "attachments must come inline with the page, not per message")
        # page count: one request per folder page, not per message
        page_calls = [url for url, _ in support.FakeRequests.calls
                      if url.endswith("/messages") or "/next/" in url]
        self.assertLessEqual(len(page_calls), 9)  # 4 folders x (page + nextLink)
        # every message carries its inline attachment with content
        missing = [(m["id"], m["attachments"]) for m in messages
                   if not m["attachments"]]
        self.assertEqual(missing, [])
        no_content = [m["id"] for m in messages
                      if not m["attachments"][0].get("contentBytes")]
        self.assertEqual(no_content, [])

    def test_no_expand_when_attachments_disabled(self):
        cfg = support.make_cfg(migrate_attachments=False)
        list(mig.iter_source_emails(cfg, "u@example.com", "token"))
        page_calls = [(url, params) for url, params in support.FakeRequests.calls
                      if url.endswith("/messages") and "/next/" not in url]
        self.assertTrue(page_calls)
        for url, params in page_calls:
            self.assertNotIn("$expand", params, f"$expand in {url}")
            self.assertNotIn("attachments", params.get("$select", ""),
                             f"attachments in $select: {url} {params}")
        # and messages carry no attachment payload at all
        messages = list(mig.iter_source_emails(cfg, "u@example.com", "token"))
        self.assertTrue(all(m["attachments"] == [] for m in messages))

    def test_graph_calls_use_keep_alive_session(self):
        with mock.patch.object(mig, "_get_session") as get_session:
            session = mock.Mock()
            session.get.return_value = support.FakeResponse({"value": []})
            get_session.return_value = session
            mig.graph_get("https://graph.test/x", "token")
            session.get.assert_called_once()
            # headers still carry the token
            _, kwargs = session.get.call_args
            self.assertEqual(kwargs["headers"]["Authorization"], "Bearer token")


class LargeAttachmentFallbackTests(unittest.TestCase):
    def setUp(self):
        support.install_fakes()
        self.cfg = support.make_cfg(migrate_attachments=True)

    def test_missing_content_bytes_fetched_via_batch(self):
        # Make every inline attachment look "large" (no contentBytes) and
        # provide the content for the $value fallback.
        support.FakeRequests.omit_content_for = {
            f"att-{i * 100}" for i in (1, 2, 3)}  # m1..m3 in all 4 folders
        for att_id in support.FakeRequests.omit_content_for:
            support.FakeRequests.attachment_content[att_id] = \
                f"big-content-{att_id}".encode()

        messages = list(mig.iter_source_emails(self.cfg, "u@example.com", "token"))
        self.assertEqual(len(messages), 12)
        missing = [(m["id"], m["attachments"]) for m in messages
                   if not m["attachments"]]
        self.assertEqual(missing, [],
                         "every message must carry its inline attachment")
        with_content = [m for m in messages
                        if m["attachments"] and m["attachments"][0].get("contentBytes")]
        self.assertEqual(len(with_content), 12,
                         [(m["id"], bool(m["attachments"])) for m in messages])

        batch_calls = [url for url, _ in support.FakeRequests.calls
                       if url.endswith("/$batch")]
        self.assertTrue(batch_calls, "expected at least one $batch call")
        # 12 messages x 1 attachment, all missing contentBytes, but the fill runs
        # per page: 2 pages x 4 folders = 8 batches of 1-2 jobs each.
        self.assertLessEqual(len(batch_calls), 8)
        # content matches what $value returned
        import base64
        sample = messages[0]["attachments"][0]
        self.assertEqual(base64.b64decode(sample["contentBytes"]),
                         f"big-content-{sample['id']}".encode())

    def test_batch_jobs_are_chunked_at_20(self):
        support.FakeRequests.message_pages = {
            "f1": [[support.message(i) for i in range(1, 46)]],  # 45 messages
            "f1c": [[support.message(100)]],
            "f2": [[support.message(101)]],
            "f3": [[support.message(102)]],
        }
        support.FakeRequests.omit_content_for = {
            f"att-{i * 100}" for i in range(1, 46)}
        for att_id in support.FakeRequests.omit_content_for:
            support.FakeRequests.attachment_content[att_id] = b"x"

        messages = list(mig.iter_source_emails(self.cfg, "u@example.com", "token"))
        self.assertEqual(len(messages), 48)
        batch_calls = [url for url, _ in support.FakeRequests.calls
                       if url.endswith("/$batch")]
        # 45 large attachments -> ceil(45/20) = 3 batches for f1
        self.assertEqual(len(batch_calls), 3)
        filled = [m for m in messages if m["attachments"][0].get("contentBytes")]
        self.assertEqual(len(filled), 48)

    def test_zero_byte_attachment_is_reported_not_fatal(self):
        # att-100 (messages m1..m3) has no contentBytes and no $value content
        support.FakeRequests.omit_content_for = {"att-100"}
        # no attachment_content entry -> $value returns the default empty bytes
        messages = list(mig.iter_source_emails(self.cfg, "u@example.com", "token"))
        self.assertEqual(len(messages), 12)
        empty = [m for m in messages
                 if m["attachments"] and not m["attachments"][0].get("contentBytes")]
        # att-100 is omitted: the m1 message of every folder (4 messages);
        # the $value fallback returns empty for it, like Graph does for the
        # broken inline images found in the real mailbox
        self.assertEqual(sorted(m["id"] for m in empty), ["m1"] * 4,
                         [(m["id"], m["attachments"]) for m in empty])
        # conversion still works; the attachment is skipped with a warning
        rfc822 = mig.convert_to_rfc822(self.cfg, empty[0])
        self.assertIsInstance(rfc822, str)
        # and a message whose attachment has content converts with it intact
        with_content = next(m for m in messages
                            if m["attachments"][0].get("contentBytes"))
        ok = mig.convert_to_rfc822(self.cfg, with_content)
        self.assertIn("Content-Disposition", ok)
        self.assertIn(with_content["attachments"][0]["name"], ok)

    def test_item_attachment_without_inline_item_is_expanded(self):
        message = support.message(1)
        message["attachments"] = [{
            "@odata.type": "#microsoft.graph.itemAttachment",
            "id": "att-item-9",
            "name": "nested.eml",
        }]
        mig.fill_attachment_content("token", "u@example.com", [message])
        self.assertIsInstance(message["attachments"][0].get("item"), dict)


class PageShrinkTests(unittest.TestCase):
    def setUp(self):
        support.install_fakes()
        self.cfg = support.make_cfg(migrate_attachments=False, graph_page_size=100)

    def test_504_shrinks_page_size_and_completes(self):
        support.FakeRequests.fail_pages_with_top_above = 50
        messages = list(mig.iter_source_emails(self.cfg, "u@example.com", "token"))
        self.assertEqual(len(messages), 12)
        page_tops = [params.get("$top") for url, params in support.FakeRequests.calls
                     if url.endswith("/messages")]
        self.assertIn(100, page_tops)   # first attempt at configured size
        self.assertIn(50, page_tops)    # shrunk after the 504

    def test_504_below_floor_is_raised(self):
        support.FakeRequests.fail_pages_with_top_above = 0  # everything 504s
        with self.assertRaises(mig.GraphError):
            list(mig.iter_source_emails(self.cfg, "u@example.com", "token"))


class TopHelpersTests(unittest.TestCase):
    def test_with_top_in_params(self):
        url, params = mig._with_top("https://graph.test/messages",
                                    {"$top": 100, "$filter": "x"}, 50)
        self.assertEqual(params["$top"], 50)
        self.assertEqual(params["$filter"], "x")

    def test_with_top_in_url(self):
        url, params = mig._with_top(
            "https://graph.test/messages?%24top=100&%24filter=x", None, 50)
        self.assertIsNone(params)
        self.assertIn("%24top=50", url)
        self.assertNotIn("%24top=100", url)
        self.assertIn("%24filter=x", url)

    def test_page_top_reads_params_then_url(self):
        self.assertEqual(mig._page_top("https://x/messages", {"$top": 100}), 100)
        self.assertEqual(
            mig._page_top("https://x/messages?%24top=250", None), 250)
        self.assertIsNone(mig._page_top("https://x/messages", None))


if __name__ == "__main__":
    unittest.main()
