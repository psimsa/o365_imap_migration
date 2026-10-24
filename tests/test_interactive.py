import io
import os
import stat
import sys
import tempfile
import unittest
from unittest import mock

from tests import support

mig = support.mig
import interactive as tui  # noqa: E402


class DummyBoard:
    def __init__(self, *args, **kwargs):
        pass

    def register(self, label, total):
        return label

    def advance(self, key, ok=True):
        pass

    def finish(self, key, status="done"):
        pass

    def set_folder(self, key, folder):
        pass

    def log(self, key, message):
        pass

    def notify(self, message):
        pass


class PureHelperTests(unittest.TestCase):
    def test_root_candidates_exclude_nested_folders(self):
        names = ["INBOX", "Archive", "Archive/2024", "Sent Items", "A/B/C"]
        self.assertEqual(tui.root_folder_candidates(names, "/"),
                         ["INBOX", "Archive", "Sent Items"])

    def test_special_roots_and_matching(self):
        self.assertTrue(tui.is_special_root("Inbox"))
        self.assertTrue(tui.is_special_root("Junk Email"))
        self.assertFalse(tui.is_special_root("Projects"))
        self.assertEqual(tui.match_folder("archive", ["Archive", "Projects"]), "Archive")
        self.assertIsNone(tui.match_folder("new", ["Archive"]))

    def test_parent_name_validation(self):
        self.assertEqual(tui.validate_parent_name(" Migration 2026 ", "/"),
                         "Migration 2026")
        with self.assertRaises(ValueError):
            tui.validate_parent_name("", "/")
        with self.assertRaises(ValueError):
            tui.validate_parent_name("a/b", "/")

    def test_build_plan_with_parent_folder(self):
        cfg = support.make_cfg(target_parent_folder="Migration")
        folders = support.sample_folders()
        plan = tui.build_plan(cfg, tui.MailboxChoice("u@example.com"), folders,
                              "/", ["INBOX", "Archive"])
        self.assertEqual(plan.create_count, 5)
        self.assertEqual(plan.reuse_count, 0)
        self.assertEqual(plan.total_messages, 8)
        self.assertFalse(plan.parent_exists)
        self.assertTrue(any("will be created" in warning for warning in plan.warnings))

    def test_build_plan_counts_existing_folders(self):
        cfg = support.make_cfg(target_parent_folder="Migration")
        folders = support.sample_folders()
        plan = tui.build_plan(cfg, tui.MailboxChoice("u@example.com"), folders, "/",
                              ["INBOX", "Migration", "Migration/Inbox"])
        self.assertEqual(plan.create_count, 3)
        self.assertEqual(plan.reuse_count, 2)
        self.assertTrue(plan.parent_exists)

    def test_build_plan_default_mapping_warns_on_fallback(self):
        cfg = support.make_cfg()
        folders = support.sample_folders() + [
            {"id": "f4", "displayName": "Junk Email", "path": "Junk Email",
             "totalItemCount": 2}]
        plan = tui.build_plan(cfg, tui.MailboxChoice("u@example.com"), folders, "/",
                              ["INBOX", "Archive", "Sent"])
        self.assertEqual(plan.create_count, 0)
        self.assertTrue(any("Junk Email" in warning for warning in plan.warnings))
        self.assertEqual(plan.total_messages, 10)

    def test_build_plan_applies_limit(self):
        cfg = support.make_cfg(migrate_limit=3)
        plan = tui.build_plan(cfg, tui.MailboxChoice("u@example.com"),
                              support.sample_folders(), "/", ["INBOX"])
        self.assertEqual(plan.total_messages, 3)

    def test_write_env_values_preserves_comments_and_order(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, ".env")
            with open(path, "w", encoding="utf-8") as handle:
                handle.write("# keep me\nSOURCE_CLIENT_ID=old\nOTHER=1\n"
                             "SOURCE_CLIENT_SECRET=\n")
            tui.write_env_values(path, {
                "SOURCE_CLIENT_ID": "new",
                "SOURCE_CLIENT_SECRET": "shh",
                "SOURCE_TENANT_ID": "tenant",
            })
            with open(path, encoding="utf-8") as handle:
                lines = handle.read().splitlines()
            self.assertEqual(lines[0], "# keep me")
            self.assertIn("SOURCE_CLIENT_ID=new", lines)
            self.assertIn("OTHER=1", lines)
            self.assertIn("SOURCE_CLIENT_SECRET=shh", lines)
            self.assertIn("SOURCE_TENANT_ID=tenant", lines)
            self.assertEqual(stat.S_IMODE(os.stat(path).st_mode), 0o600)

    def test_append_csv_row_writes_header_and_dedupes(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "details.csv")
            row = ("a@x.com", "imap.y.com", "b@y.com", "pw")
            self.assertTrue(tui.append_csv_row(path, row))
            self.assertFalse(tui.append_csv_row(path, row))
            self.assertTrue(tui.append_csv_row(
                path, ("c@x.com", "imap.y.com", "d@y.com", "pw2")))
            with open(path, encoding="utf-8") as handle:
                lines = [line.strip() for line in handle if line.strip()]
            self.assertEqual(lines[0], ",".join(mig.CSV_HEADER))
            self.assertEqual(len(lines), 3)
            self.assertEqual(stat.S_IMODE(os.stat(path).st_mode), 0o600)

    def test_plan_shows_time_filter_and_resume(self):
        cfg = support.make_cfg(restartable=True,
                               migrate_after="2020-01-01T00:00:00Z")
        plan = tui.build_plan(cfg, tui.MailboxChoice("u@example.com"),
                              support.sample_folders(), "/", ["INBOX"],
                              resume_count=4)
        rows = dict(tui.plan_rows(plan))
        self.assertIn("after 2020-01-01", rows["Time filter"])
        self.assertIn("4 message(s)", rows["Restartable"])

    def test_scripted_confirm_defaults(self):
        self.assertFalse(tui.ScriptedPrompter([None]).confirm("q", default=False))
        self.assertTrue(tui.ScriptedPrompter([None]).confirm("q", default=True))
        self.assertTrue(tui.ScriptedPrompter(["yes"]).confirm("q"))
        self.assertFalse(tui.ScriptedPrompter(["n"]).confirm("q", default=True))


class FlowTests(unittest.TestCase):
    def setUp(self):
        support.install_fakes()
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.cfg = support.make_cfg(
            client_id="id", client_secret="secret", tenant_id="tenant",
            csv_file=os.path.join(self.tmp.name, "details.csv"))
        self.args = mig.parse_args(["-i"])
        self.renderer = tui.PlainRenderer()

    def _patch_world(self, graph_get=None):
        self.mail = support.FakeIMAP()
        patches = [
            mock.patch.object(mig, "load_config", return_value=self.cfg),
            mock.patch.object(tui, "verify_credentials", return_value=None),
            mock.patch.object(mig, "authenticate_account",
                              return_value=("token", None)),
            mock.patch.object(mig, "get_mail_folders",
                              return_value=support.sample_folders()),
            mock.patch.object(mig, "connect_to_target_imap", return_value=self.mail),
            mock.patch.object(mig, "get_imap_hierarchy",
                              return_value=("/", ["INBOX", "Archive"])),
            mock.patch.object(mig, "ProgressBoard", DummyBoard),
            mock.patch.object(mig, "migrate_mailbox", return_value=(2, 0)),
        ]
        if graph_get is None:
            patches.append(mock.patch.object(
                mig, "graph_get", return_value={"value": [support.message(1)]}))
        else:
            patches.append(mock.patch.object(mig, "graph_get", side_effect=graph_get))

        started = [patch.start() for patch in patches]
        for patch in patches:
            self.addCleanup(patch.stop)
        return started

    def _answers(self):
        return [
            "user@contoso.com",   # source mailbox
            "imap.example.com",   # IMAP server
            "target@example.com",  # target mailbox
            "pw",                 # IMAP password
            False,                # save row to CSV
            False,                # mirror under parent folder
            False,                # do not resume an interrupted migration
            0,                    # all folders
            False,                # no message limit
            False,                # no received-date filter
        ]

    def test_full_flow_migrates_after_confirmation(self):
        self._patch_world()
        prompter = tui.ScriptedPrompter(self._answers() + [True])
        rc = tui.run_interactive(self.args, prompter, self.renderer)

        self.assertEqual(rc, 0, self.renderer.lines)
        self.assertTrue(any("Start the migration" in question
                            for kind, question in prompter.prompts))
        self.assertTrue(any("Validation & plan" in line
                            for line in self.renderer.lines))
        self.assertTrue(any("Sample message" in line and "Inbox" in line
                            for line in self.renderer.lines))
        migrate_mock = mig.migrate_mailbox
        self.assertEqual(migrate_mock.call_count, 1)
        mailbox = migrate_mock.call_args.args[1]
        self.assertEqual(mailbox, ("user@contoso.com", "imap.example.com",
                                   "target@example.com", "pw"))
        self.assertTrue(any("Rerun non-interactively" in line
                            for line in self.renderer.lines))
        self.assertTrue(any("Time filter" in line for line in self.renderer.lines))
        self.assertTrue(any("Restartable" in line for line in self.renderer.lines))
        cfg_used = mig.migrate_mailbox.call_args.args[0]
        self.assertFalse(cfg_used.restartable)

    def test_date_filter_is_applied(self):
        self._patch_world()
        answers = self._answers()[:-1] + [True, "2020-01-01", "2021-01-01", True]
        prompter = tui.ScriptedPrompter(answers)
        rc = tui.run_interactive(self.args, prompter, self.renderer)
        self.assertEqual(rc, 0, self.renderer.lines)
        cfg_used = mig.migrate_mailbox.call_args.args[0]
        self.assertEqual(cfg_used.migrate_after, "2020-01-01T00:00:00Z")
        self.assertEqual(cfg_used.migrate_before, "2021-01-01T00:00:00Z")
        self.assertTrue(any("after 2020-01-01" in line
                            for line in self.renderer.lines))

    def test_resume_prompt_enables_restartable(self):
        self._patch_world()
        answers = list(self._answers())
        answers[6] = True  # resume if interrupted
        answers.append(True)  # start
        prompter = tui.ScriptedPrompter(answers)
        rc = tui.run_interactive(self.args, prompter, self.renderer)
        self.assertEqual(rc, 0, self.renderer.lines)
        cfg_used = mig.migrate_mailbox.call_args.args[0]
        self.assertTrue(cfg_used.restartable)
        self.assertFalse(cfg_used.reset_session)
        self.assertTrue(any("Progress will be recorded" in line
                            for line in self.renderer.lines))

    def test_resume_prompt_offers_existing_session(self):
        self.cfg = support.make_cfg(
            client_id="id", client_secret="secret", tenant_id="tenant",
            csv_file=os.path.join(self.tmp.name, "details.csv"),
            session_dir=self.tmp.name)
        mailbox = ("user@contoso.com", "imap.example.com",
                   "target@example.com", "pw")
        seed = mig.MigrationSession.load(self.cfg, mailbox)
        seed.record("m1", "Inbox")
        seed.close()

        self._patch_world()
        answers = ["user@contoso.com", "imap.example.com", "target@example.com",
                   "pw", False, False, True, False, 0, False, False, True]
        prompter = tui.ScriptedPrompter(answers)
        rc = tui.run_interactive(self.args, prompter, self.renderer)
        self.assertEqual(rc, 0, self.renderer.lines)
        cfg_used = mig.migrate_mailbox.call_args.args[0]
        self.assertTrue(cfg_used.restartable)
        self.assertFalse(cfg_used.reset_session)
        self.assertTrue(any("1 message(s) already migrated" in line
                            for line in self.renderer.lines))

    def test_confirmation_defaults_to_no(self):
        self._patch_world()
        prompter = tui.ScriptedPrompter(self._answers() + [False])
        rc = tui.run_interactive(self.args, prompter, self.renderer)
        self.assertEqual(rc, 2)
        self.assertEqual(mig.migrate_mailbox.call_count, 0)
        self.assertTrue(any("nothing written" in line
                            for line in self.renderer.lines))

    def test_failed_validation_blocks_migration(self):
        self._patch_world(graph_get=RuntimeError("graph exploded"))
        prompter = tui.ScriptedPrompter(self._answers() + [False])
        rc = tui.run_interactive(self.args, prompter, self.renderer)
        self.assertEqual(rc, 2)
        self.assertEqual(mig.migrate_mailbox.call_count, 0)
        self.assertTrue(any("Validation failed" in question
                            for kind, question in prompter.prompts))

    def test_dry_run_flag_is_ignored_with_note(self):
        self._patch_world()
        with mock.patch.object(tui.Wizard, "run", return_value=0):
            args = mig.parse_args(["-i", "--dry-run"])
            rc = tui.run_interactive(args, tui.ScriptedPrompter([]), self.renderer)
        self.assertEqual(rc, 0)
        self.assertTrue(any("--dry-run is ignored" in line
                            for line in self.renderer.lines))

    def test_non_tty_is_rejected(self):
        with mock.patch.object(sys, "stdin", io.StringIO()), \
                mock.patch.object(sys, "stdout", io.StringIO()):
            rc = tui.run_interactive(self.args)
        self.assertEqual(rc, 2)


class CredentialFlowTests(unittest.TestCase):
    def test_retries_then_succeeds(self):
        cfg = support.make_cfg(client_id="bad", client_secret="bad", tenant_id="bad")
        renderer = tui.PlainRenderer()
        prompter = tui.ScriptedPrompter(["id1", "secret1", "tenant1", False])
        with mock.patch.object(tui, "verify_credentials",
                               side_effect=["rejected", None]):
            result = tui.credentials_flow(cfg, prompter, renderer,
                                          "/tmp/does-not-exist.env")
        self.assertEqual(result.client_id, "id1")
        self.assertTrue(any("verified" in line for line in renderer.lines))

    def test_gives_up_after_three_attempts(self):
        cfg = support.make_cfg(client_id="bad", client_secret="bad", tenant_id="bad")
        renderer = tui.PlainRenderer()
        prompter = tui.ScriptedPrompter(["id1", "s1", "t1"])
        with mock.patch.object(tui, "verify_credentials", return_value="rejected"):
            with self.assertRaises(tui.InteractiveError):
                tui.credentials_flow(cfg, prompter, renderer,
                                     "/tmp/does-not-exist.env")


if __name__ == "__main__":
    unittest.main()
