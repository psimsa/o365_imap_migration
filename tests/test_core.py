import json
import os
import ssl
import tempfile
import unittest
from unittest import mock

from tests import support

mig = support.mig


def stat_mode(path):
    return os.stat(path).st_mode & 0o777


class CoreTests(unittest.TestCase):
    def setUp(self):
        support.install_fakes()
        self.cfg = support.make_cfg()

    def test_follows_next_links_for_folders_and_messages(self):
        folders = mig.get_mail_folders(self.cfg, "token", "u@example.com")
        self.assertEqual({folder["path"] for folder in folders},
                         {"Inbox", "Inbox/Sub", "Archive", "Sent Items"})
        messages = mig.fetch_all_emails(self.cfg, "u@example.com", "token")
        self.assertEqual(len(messages), 12)  # 3 messages x 4 folders
        self.assertTrue(all(item["folderPath"] for item in messages))

    def test_message_limit_is_applied(self):
        cfg = support.make_cfg(migrate_limit=5)
        messages = mig.fetch_all_emails(cfg, "u@example.com", "token")
        self.assertEqual(len(messages), 5)

    def test_folder_filter_selects_subtree(self):
        folders = mig.get_mail_folders(self.cfg, "token", "u@example.com")
        selected, missing = mig.select_folders(folders, ["Inbox"])
        self.assertEqual({folder["path"] for folder in selected},
                         {"Inbox", "Inbox/Sub"})
        self.assertEqual(missing, [])

    def test_default_mapping_falls_back_to_inbox(self):
        class NoSpamIMAP(support.FakeIMAP):
            def list(self):
                return ("OK", [b'(\\HasNoChildren) "/" "INBOX"',
                               b'(\\HasNoChildren) "/" "Drafts"',
                               b'(\\HasNoChildren) "/" "Sent"',
                               b'(\\HasNoChildren) "/" "Deleted Items"'])

        target = NoSpamIMAP()
        messages = [
            dict(support.message(1), folderName="Inbox", folderPath="Inbox"),
            dict(support.message(2), folderName="Junk Email", folderPath="Junk Email"),
        ]
        migrated, failed = mig.migrate_emails(self.cfg, target, messages)
        self.assertEqual((migrated, failed), (2, 0))
        self.assertEqual([name for name, _ in target.appended], ["INBOX", "INBOX"])
        self.assertEqual(target.created, [])

    def test_parent_folder_is_created_and_used(self):
        cfg = support.make_cfg(target_parent_folder="Migration")
        target = support.FakeIMAP()
        messages = mig.fetch_all_emails(self.cfg, "u@example.com", "token")
        migrated, failed = mig.migrate_emails(cfg, target, messages)
        self.assertEqual((migrated, failed), (12, 0))
        self.assertEqual(set(target.created), {
            "Migration", "Migration/Inbox", "Migration/Inbox/Sub",
            "Migration/Archive", "Migration/Sent Items"})
        self.assertEqual(target.created.index("Migration"), 0)
        self.assertEqual({name for name, _ in target.appended},
                         {"Migration/Inbox", "Migration/Inbox/Sub",
                          "Migration/Archive", "Migration/Sent Items"})

    def test_tls_verification_default_and_opt_out(self):
        mail = mig.connect_to_target_imap(self.cfg, "imap.example.com", "u", "p")
        self.assertEqual(mail.ssl_context.verify_mode, ssl.CERT_REQUIRED)
        self.assertTrue(mail.ssl_context.check_hostname)

        insecure = mig.connect_to_target_imap(
            support.make_cfg(imap_ssl_verify=False), "imap.example.com", "u", "p")
        self.assertEqual(insecure.ssl_context.verify_mode, ssl.CERT_NONE)

    def test_load_config_can_skip_credential_check(self):
        with tempfile.TemporaryDirectory() as tmp:
            # Isolate from the developer's real .env file.
            env_path = os.path.join(tmp, ".env")
            with open(env_path, "w", encoding="utf-8"):
                pass
            with mock.patch.dict(os.environ, {}, clear=True):
                cfg = mig.load_config(env_file=env_path, require_credentials=False)
                self.assertEqual(cfg.client_id, "")
                self.assertTrue(cfg.migrate_attachments)
                with self.assertRaises(mig.MigrationError):
                    mig.load_config(env_file=env_path, require_credentials=True)

    def test_migrate_mailbox_returns_counts(self):
        messages = [support.message(1), support.message(2)]
        with mock.patch.object(mig, "authenticate_account",
                               return_value=("token", None)), \
                mock.patch.object(mig, "fetch_all_emails", return_value=messages), \
                mock.patch.object(mig, "connect_to_target_imap",
                                  return_value=support.FakeIMAP()):
            migrated, failed = mig.migrate_mailbox(
                self.cfg, ("u@example.com", "imap.example.com", "t@example.com", "pw"))
        self.assertEqual((migrated, failed), (2, 0))

    def test_defaults_enable_attachments_and_html(self):
        cfg = mig.Config(client_id="", client_secret="", tenant_id="")
        self.assertTrue(cfg.migrate_attachments)
        self.assertEqual(cfg.email_format, "html")

    def test_date_parsing_and_graph_filter(self):
        self.assertEqual(mig.normalize_date_bound("2024-03-01"),
                         "2024-03-01T00:00:00Z")
        self.assertEqual(mig.normalize_date_bound("2024-03-01T12:30:00+02:00"),
                         "2024-03-01T10:30:00Z")
        cfg = support.make_cfg(migrate_after="2020-01-01",
                               migrate_before="2021-06-01")
        self.assertEqual(
            mig.build_received_filter(cfg),
            "receivedDateTime ge 2020-01-01T00:00:00Z and "
            "receivedDateTime lt 2021-06-01T00:00:00Z")
        with self.assertRaises(mig.MigrationError):
            mig.normalize_date_bound("not-a-date")

    def test_fetch_passes_time_filter_to_graph(self):
        cfg = support.make_cfg(migrate_after="2023-01-01",
                               migrate_before="2024-01-01")
        mig.fetch_all_emails(cfg, "u@example.com", "token")
        filters = [params.get("$filter") for url, params in support.FakeRequests.calls
                   if url.endswith("/messages")]
        self.assertTrue(filters)
        self.assertIn("receivedDateTime ge 2023-01-01T00:00:00Z", filters[0])
        self.assertIn("receivedDateTime lt 2024-01-01T00:00:00Z", filters[0])

    def test_skip_ids_applied_before_limit(self):
        messages = mig.graph_get_all(
            "https://graph.test/messages", "token", params={"$top": 100},
            limit=1, skip_ids={"m1", "m2"})
        self.assertEqual([item["id"] for item in messages], ["m3"])

    def test_session_roundtrip_reset_and_mismatch(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = support.make_cfg(restartable=True, session_dir=tmp)
            mailbox = ("u@example.com", "imap.example.com", "t@example.com", "pw")
            session = mig.MigrationSession.load(cfg, mailbox)
            session.record("m1", "Inbox")
            session.record("m2", "Inbox")
            session.close()

            path = mig.session_path(cfg, mailbox)
            self.assertTrue(os.path.exists(path))
            self.assertEqual(stat_mode(path), 0o600)
            reloaded = mig.MigrationSession.load(cfg, mailbox)
            self.assertEqual(reloaded.ids, {"m1", "m2"})
            reloaded.close()

            other = ("u@example.com", "imap.example.com", "x@example.com", "pw")
            other_path = mig.session_path(cfg, other)
            with open(other_path, "w", encoding="utf-8") as handle:
                handle.write(json.dumps({
                    "type": "header", "version": 1, "source": "u@example.com",
                    "target": "t@example.com", "server": "imap.example.com",
                    "parent_folder": ""}) + "\n")
            with self.assertRaises(mig.MigrationError):
                mig.MigrationSession.load(cfg, other)

            mig.MigrationSession.load(cfg, mailbox).reset()
            self.assertFalse(os.path.exists(path))

    def test_migrate_emails_records_only_successes(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = support.make_cfg(restartable=True, session_dir=tmp)
            mailbox = ("u@example.com", "imap.example.com", "t@example.com", "pw")
            session = mig.MigrationSession.load(cfg, mailbox)
            messages = [support.message(1), support.message(2)]
            migrated, failed = mig.migrate_emails(
                cfg, support.FakeIMAP(), messages, session=session)
            self.assertEqual((migrated, failed), (2, 0))
            self.assertEqual(session.ids, {"m1", "m2"})
            session.close()

    def test_migrate_mailbox_resumes_from_session(self):
        with tempfile.TemporaryDirectory() as tmp:
            cfg = support.make_cfg(restartable=True, session_dir=tmp)
            mailbox = ("u@example.com", "imap.example.com", "t@example.com", "pw")
            seed = mig.MigrationSession.load(cfg, mailbox)
            seed.record("m1", "Inbox")
            seed.close()

            seen = {}

            def fake_fetch(cfg_, source, token, skip_ids=None):
                seen["skip"] = set(skip_ids or ())
                return [support.message(2)]

            with mock.patch.object(mig, "authenticate_account",
                                   return_value=("token", None)), \
                    mock.patch.object(mig, "fetch_all_emails", side_effect=fake_fetch), \
                    mock.patch.object(mig, "connect_to_target_imap",
                                      return_value=support.FakeIMAP()):
                migrated, failed = mig.migrate_mailbox(cfg, mailbox)

            self.assertEqual((migrated, failed), (1, 0))
            self.assertEqual(seen["skip"], {"m1"})
            session = mig.MigrationSession.load(cfg, mailbox)
            self.assertEqual(session.ids, {"m1", "m2"})
            session.close()

    def test_utf7_and_target_paths(self):
        self.assertEqual(mig.imap_utf7_encode("~peter/mail/台北/日本語"),
                         "~peter/mail/&U,BTFw-/&ZeVnLIqe-")
        self.assertEqual(
            mig.build_target_path(support.make_cfg(target_parent_folder="Mig/2026"),
                                  "Inbox/Sub", "/"),
            "Mig/2026/Inbox/Sub")

    def test_progress_board_non_tty(self):
        class Stream:
            def __init__(self):
                self.text = ""

            def write(self, value):
                self.text += value

            def flush(self):
                pass

            def isatty(self):
                return False

        stream = Stream()
        board = mig.ProgressBoard(stream=stream, tty=False, interval=0)
        key = board.register("a@b.com", 2)
        board.advance(key, True)
        board.advance(key, False)
        board.finish(key, "done")
        self.assertIn("[a@b.com] 2/2", stream.text)
        self.assertIn("1 ok, 1 failed", stream.text)


if __name__ == "__main__":
    unittest.main()
