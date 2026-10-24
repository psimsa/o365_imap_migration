"""Tests for session records (uid/sha256/subject) and --verify."""

import os
import tempfile
import unittest
from unittest import mock

from tests import support

mig = support.mig


class SessionRecordTests(unittest.TestCase):
    def setUp(self):
        support.install_fakes()
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.cfg = support.make_cfg(restartable=True, session_dir=self.tmp.name)
        self.mailbox = ("u@example.com", "imap.example.com", "t@example.com", "pw")

    def test_record_stores_verification_fields(self):
        session = mig.MigrationSession.load(self.cfg, self.mailbox)
        session.record("m1", folder="INBOX", uid="7", uidvalidity="1",
                       sha256="abc", subject="Předmět 🤝")
        session.close()

        reloaded = mig.MigrationSession.load(self.cfg, self.mailbox)
        self.assertEqual(reloaded.entries[0]["uid"], "7")
        self.assertEqual(reloaded.entries[0]["uidvalidity"], "1")
        self.assertEqual(reloaded.entries[0]["sha256"], "abc")
        self.assertEqual(reloaded.entries[0]["subject"], "Předmět 🤝")
        self.assertEqual(reloaded.entries[0]["folder"], "INBOX")
        reloaded.close()

    def test_migration_records_uid_and_hash(self):
        support.FakeRequests.message_pages = {
            "f1": [[support.message(1)]],
            "f1c": [[support.message(2)]],
            "f2": [[support.message(3)]],
            "f3": [[support.message(4)]],
        }
        with mock.patch.object(mig, "authenticate_account",
                               return_value=("token", None)), \
                mock.patch.object(mig, "get_mail_folders",
                                  return_value=support.sample_folders()), \
                mock.patch.object(mig, "connect_to_target_imap",
                                  return_value=support.FakeIMAP()):
            migrated, failed = mig.migrate_mailbox(self.cfg, self.mailbox)

        self.assertEqual((migrated, failed), (4, 0))
        session = mig.MigrationSession.load(self.cfg, self.mailbox)
        self.assertEqual(len(session.entries), 4)
        for entry in session.entries:
            self.assertTrue(entry["uid"])
            self.assertTrue(entry["sha256"])
            self.assertTrue(entry["subject"])
            self.assertTrue(entry["folder"])
        session.close()


def unicode_message(index, subject=None):
    """A message with non-ASCII subject/attachment, like the real mailbox has."""
    msg = support.message(index)
    msg["subject"] = subject or f"Zpráva číslo {index} – příloha 🤝 中文"
    return msg


class VerifyMailboxTests(unittest.TestCase):
    """Verification over messages with non-ASCII subjects and attachments."""

    def setUp(self):
        support.install_fakes()
        # One page of 3 unique messages per folder, so UID/hash verification is
        # exercised over distinct content.
        support.FakeRequests.message_pages = {
            "f1": [[unicode_message(1), unicode_message(2), unicode_message(3)]],
            "f1c": [[unicode_message(4)]],
            "f2": [[unicode_message(5)]],
            "f3": [[unicode_message(6)]],
        }
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.cfg = support.make_cfg(restartable=True, session_dir=self.tmp.name,
                                    migrate_attachments=True)
        self.mailbox = ("u@example.com", "imap.example.com", "t@example.com", "pw")
        self.target = support.FakeIMAP()

    def _migrate(self):
        with mock.patch.object(mig, "authenticate_account",
                               return_value=("token", None)), \
                mock.patch.object(mig, "get_mail_folders",
                                  return_value=support.sample_folders()), \
                mock.patch.object(mig, "connect_to_target_imap",
                                  return_value=self.target):
            return mig.migrate_mailbox(self.cfg, self.mailbox)

    def _verify(self):
        with mock.patch.object(mig, "connect_to_target_imap",
                               return_value=self.target):
            return mig.verify_mailbox(self.cfg, self.mailbox)

    def test_all_messages_verify(self):
        migrated, failed = self._migrate()
        self.assertEqual((migrated, failed), (6, 0))
        report = self._verify()
        self.assertEqual(report.total, 6)
        self.assertEqual(report.verified, 6)
        self.assertEqual(report.subject_mismatches, 0)
        self.assertTrue(report.ok, report.problems)
        # Default mapping: Inbox+Inbox/Sub -> INBOX, Archive -> INBOX (no such
        # folder in the fake), Sent Items -> Sent.
        self.assertEqual(set(report.folders), {"INBOX", "Sent"})

    def test_missing_message_is_reported(self):
        self._migrate()
        # Remove one stored message from the "server".
        key = next(iter(self.target._stored))
        del self.target._stored[key]
        report = self._verify()
        self.assertEqual(report.missing, 1)
        self.assertFalse(report.ok)
        self.assertTrue(any("not found" in p for p in report.problems))

    def test_content_tampering_is_detected(self):
        self._migrate()
        # Tamper with a message whose subject survives, so only the hash differs.
        key = next(iter(self.target._stored))
        self.target._stored[key] = self.target._stored[key] + b"X"
        report = self._verify()
        self.assertEqual(report.mismatched, 1)
        self.assertFalse(report.ok)

    def test_subject_decode_mismatch_is_detected(self):
        self._migrate()
        session = mig.MigrationSession.load(self.cfg, self.mailbox)
        session.entries[0]["subject"] = "něco úplně jiného"
        session.entries[0]["_tampered"] = True
        session.close()
        # Rewrite the session file with the tampered subject.
        import json
        path = mig.session_path(self.cfg, self.mailbox)
        with open(path, encoding="utf-8") as handle:
            lines = handle.read().splitlines()
        with open(path, "w", encoding="utf-8") as handle:
            for line in lines:
                entry = json.loads(line)
                if entry.get("id") == session.entries[0]["id"]:
                    entry["subject"] = "něco úplně jiného"
                handle.write(json.dumps(entry) + "\n")

        report = self._verify()
        self.assertEqual(report.subject_mismatches, 1)
        self.assertFalse(report.ok)

    def test_legacy_entries_are_unverifiable_not_failed(self):
        self._migrate()
        import json
        path = mig.session_path(self.cfg, self.mailbox)
        with open(path, encoding="utf-8") as handle:
            lines = handle.read().splitlines()
        with open(path, "w", encoding="utf-8") as handle:
            for line in lines:
                entry = json.loads(line)
                if entry.get("type") != "header":
                    entry.pop("uid", None)
                    entry.pop("sha256", None)
                handle.write(json.dumps(entry) + "\n")

        report = self._verify()
        self.assertEqual(report.unverifiable, 6)
        self.assertEqual(report.verified, 0)
        self.assertTrue(report.ok)  # unverifiable is not a failure

    def test_uidvalidity_change_is_unverifiable(self):
        self._migrate()
        self.target.uidvalidity = 999  # mailbox was rebuilt on the server
        report = self._verify()
        self.assertEqual(report.unverifiable, 6)
        self.assertTrue(any("UIDVALIDITY changed" in p for p in report.problems))

    def test_no_session_means_empty_report(self):
        report = self._verify()
        self.assertEqual(report.total, 0)
        self.assertTrue(report.ok)

    def test_line_ending_normalization_is_tolerated(self):
        self._migrate()
        # Simulate a server that returns LF endings and a trailing CRLF.
        for key, payload in list(self.target._stored.items()):
            self.target._stored[key] = payload.replace(b"\r\n", b"\n") + b"\r\n"
        report = self._verify()
        self.assertEqual(report.verified, 6)
        self.assertTrue(report.ok, report.problems)

    def test_format_verification_lines(self):
        self._migrate()
        report = self._verify()
        lines = mig.format_verification(report)
        self.assertTrue(any("entries: 6" in line for line in lines))
        self.assertTrue(any("folder INBOX" in line for line in lines))


if __name__ == "__main__":
    unittest.main()
