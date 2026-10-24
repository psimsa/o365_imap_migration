import os
import ssl
import tempfile
import unittest
from unittest import mock

from tests import support

mig = support.mig


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
