"""Interactive single-mailbox migration wizard (draft).

Used by ``o365_imap_migration.py --interactive``. All Microsoft Graph and IMAP
work is delegated to the core module; this file only asks questions, renders
previews and orchestrates the existing functions.

The wizard never writes before the explicit confirmation. Its final screen is a
read-only validation (Graph auth, mailbox access, one sample message, IMAP login,
folder plan) and the confirmation defaults to "no".
"""

from __future__ import annotations

import csv
import os
import shlex
import shutil
import sys
import time
from dataclasses import dataclass, field, replace
from typing import Sequence

import o365_imap_migration as mig

TUI_HINT = ("Interactive mode needs extra packages: "
            "python -m pip install -r requirements-tui.txt")
ENV_KEYS = ("SOURCE_CLIENT_ID", "SOURCE_CLIENT_SECRET", "SOURCE_TENANT_ID")
SPECIAL_ROOTS = {
    "inbox", "drafts", "sent", "sent items", "deleted items", "trash",
    "junk", "junk email", "spam", "outbox", "archive",
}


class InteractiveError(Exception):
    """A wizard step cannot continue."""


# Prompts --------------------------------------------------------------------

class ScriptedPrompter:
    """Test double: replays answers and records every prompt."""

    def __init__(self, answers: Sequence):
        self.answers = list(answers)
        self.prompts: list[tuple[str, str]] = []

    def _answer(self, kind, question):
        if not self.answers:
            raise AssertionError(f"no scripted answer left for {kind} prompt: {question!r}")
        self.prompts.append((kind, question))
        return self.answers.pop(0)

    def text(self, question, default="", validate=None, completer=None):
        while True:
            value = self._answer("text", question)
            value = default if value is None else str(value)
            if validate is not None:
                try:
                    validate(value)
                except ValueError:
                    continue
            return value.strip()

    def secret(self, question, default=""):
        return str(self._answer("secret", question))

    def select(self, question, options, default=0):
        value = self._answer("select", question)
        if isinstance(value, int):
            return value
        if value in (None, ""):
            return default
        text = str(value)
        return int(text) - 1 if text.isdigit() else options.index(text)

    def multi_select(self, question, options, preselected=()):
        value = self._answer("multi_select", question)
        if isinstance(value, (list, tuple)):
            return [item if isinstance(item, int) else options.index(item) for item in value]
        if value in (None, ""):
            return list(preselected)
        text = str(value)
        if text.lower() == "all":
            return list(range(len(options)))
        return [int(part.strip()) - 1 for part in text.split(",") if part.strip()]

    def confirm(self, question, default=False):
        value = self._answer("confirm", question)
        if value is None:
            return default
        if isinstance(value, bool):
            return value
        return str(value).strip().lower() in {"y", "yes", "true", "1"}


class PromptToolkitPrompter:
    """Real prompts: hidden secrets and fuzzy typeahead for folder names."""

    def __init__(self):
        try:
            import prompt_toolkit
            from prompt_toolkit.completion import FuzzyCompleter, WordCompleter
        except ImportError as exc:  # pragma: no cover - dependency check
            raise InteractiveError(TUI_HINT) from exc
        self._completion = (FuzzyCompleter, WordCompleter)

    def _ask(self, question, default="", is_password=False, completer=None):
        from prompt_toolkit import prompt

        kwargs = {"default": default, "is_password": is_password}
        if completer:
            fuzzy, word = self._completion
            kwargs["completer"] = fuzzy(word(list(completer), ignore_case=True))
            kwargs["complete_while_typing"] = True
        return prompt(f"{question} ", **kwargs)

    def text(self, question, default="", validate=None, completer=None):
        while True:
            value = self._ask(question, default=default, completer=completer).strip()
            if validate is not None:
                try:
                    validate(value)
                except ValueError as exc:
                    print(f"  {exc}")
                    continue
            return value

    def secret(self, question, default=""):
        return self._ask(question, default=default, is_password=True).strip()

    def select(self, question, options, default=0):
        print(question)
        for index, option in enumerate(options):
            marker = ">" if index == default else " "
            print(f" {marker} {index + 1}) {option}")

        def validate(value):
            if not value.isdigit() or not 1 <= int(value) <= len(options):
                raise ValueError(f"enter a number between 1 and {len(options)}")

        answer = self.text(f"Select [{default + 1}]", default=str(default + 1),
                           validate=validate)
        return int(answer) - 1

    def multi_select(self, question, options, preselected=()):
        print(question)
        for index, option in enumerate(options):
            marker = "*" if index in preselected else " "
            print(f" {marker} {index + 1}) {option}")

        def validate(value):
            if value.lower() in ("", "all"):
                return
            for part in value.split(","):
                part = part.strip()
                if not part.isdigit() or not 1 <= int(part) <= len(options):
                    raise ValueError(f"invalid selection {part!r}")

        answer = self.text("Select numbers separated by commas, or 'all'", validate=validate)
        if answer.lower() == "all":
            return list(range(len(options)))
        if not answer:
            return list(preselected)
        return [int(part.strip()) - 1 for part in answer.split(",") if part.strip()]

    def confirm(self, question, default=False):
        suffix = "Y/n" if default else "y/N"
        answer = self.text(f"{question} [{suffix}]", default="").strip().lower()
        if not answer:
            return default
        return answer in ("y", "yes")


# Rendering ------------------------------------------------------------------

class PlainRenderer:
    """Plain-text renderer (tests and minimal terminals)."""

    def __init__(self):
        self.lines: list[str] = []

    def _print(self, message=""):
        self.lines.append(message)
        print(message)

    def step(self, number, total, title):
        self._print(f"\nStep {number}/{total}  {title}")

    def info(self, message):
        self._print(message)

    def warn(self, message):
        self._print(f"WARNING: {message}")

    def error(self, message):
        self._print(f"ERROR: {message}")

    def stats(self, folders):
        total = sum(folder.get("totalItemCount") or 0 for folder in folders)
        self._print("Account statistics")
        ordered = sorted(folders, key=lambda f: f.get("totalItemCount") or 0, reverse=True)
        for folder in ordered[:15]:
            size = folder.get("sizeInBytes")
            size_text = f"{size / (1024 ** 2):.1f} MB" if size else "-"
            self._print(f"  {folder['path']:<40} {folder.get('totalItemCount') or 0:>8}  "
                        f"{size_text:>10}")
        if len(ordered) > 15:
            self._print(f"  ... {len(ordered) - 15} more folder(s)")
        self._print(f"  Total: {len(folders)} folder(s), {total} message(s)")

    def plan(self, checks, plan):
        self._print("\nValidation & plan")
        for check in checks:
            status = "OK  " if check.ok else "FAIL"
            self._print(f"  {status} {check.label:<24} {check.detail}")
        self._print("")
        for label, value in plan_rows(plan):
            self._print(f"  {label:<15} {value}")
        for warning in plan.warnings:
            self._print(f"  WARNING: {warning}")


class RichRenderer:
    """Rich-based renderer for a real terminal."""

    def __init__(self):
        try:
            from rich.console import Console
            from rich.panel import Panel
            from rich.table import Table
        except ImportError as exc:  # pragma: no cover - dependency check
            raise InteractiveError(TUI_HINT) from exc
        self.console = Console()
        self.table = Table
        self.panel = Panel

    def step(self, number, total, title):
        self.console.rule(f"Step {number}/{total}  {title}")

    def info(self, message):
        self.console.print(message)

    def warn(self, message):
        self.console.print(f"[yellow]WARNING:[/] {message}")

    def error(self, message):
        self.console.print(f"[red]ERROR:[/] {message}")

    def stats(self, folders):
        table = self.table(title="Account statistics", header_style="bold")
        table.add_column("Folder")
        table.add_column("Messages", justify="right")
        table.add_column("Size", justify="right")
        ordered = sorted(folders, key=lambda f: f.get("totalItemCount") or 0, reverse=True)
        for folder in ordered[:15]:
            size = folder.get("sizeInBytes")
            size_text = f"{size / (1024 ** 2):.1f} MB" if size else "-"
            table.add_row(folder["path"], str(folder.get("totalItemCount") or 0), size_text)
        if len(ordered) > 15:
            table.add_row(f"... {len(ordered) - 15} more folder(s)", "", "")
        total = sum(folder.get("totalItemCount") or 0 for folder in folders)
        table.add_row("Total", str(total), "")
        self.console.print(table)

    def plan(self, checks, plan):
        table = self.table(title="Validation & plan", header_style="bold")
        table.add_column("Check")
        table.add_column("Result")
        table.add_column("Detail")
        for check in checks:
            result = "[green]OK[/]" if check.ok else "[red]FAIL[/]"
            table.add_row(check.label, result, check.detail)
        self.console.print(table)
        for label, value in plan_rows(plan):
            self.console.print(f"[bold]{label:<15}[/] {value}")
        for warning in plan.warnings:
            self.console.print(f"[yellow]WARNING:[/] {warning}")


# Data -----------------------------------------------------------------------

@dataclass
class Check:
    label: str
    ok: bool
    detail: str = ""


@dataclass
class Plan:
    source_email: str
    target_email: str
    target_server: str
    folders: list[str]
    total_messages: int
    limit: int
    parent_folder: str
    parent_exists: bool
    create_count: int
    reuse_count: int
    attachments: bool
    email_format: str
    imap_ssl_verify: bool
    warnings: list[str] = field(default_factory=list)


@dataclass
class MailboxChoice:
    source_email: str
    target_server: str = ""
    target_email: str = ""
    password: str = ""
    from_csv: bool = False


# Pure helpers ---------------------------------------------------------------

def validate_email(value):
    text = (value or "").strip()
    if "@" not in text or " " in text:
        raise ValueError("enter a valid email address")
    return text


def validate_positive_int(value):
    if not str(value).strip().isdigit() or int(value) <= 0:
        raise ValueError("enter a positive integer")
    return int(value)


def validate_parent_name(value, delimiter="/"):
    text = (value or "").strip()
    if not text:
        raise ValueError("enter a folder name")
    if delimiter and delimiter in text:
        raise ValueError(f"enter a single root folder (no {delimiter!r})")
    return text


def root_folder_candidates(names, delimiter):
    """Top-level mailboxes only (no hierarchy delimiter in the name)."""
    return [name for name in names if not delimiter or delimiter not in name]


def is_special_root(name):
    return name.strip().lower() in SPECIAL_ROOTS


def match_folder(value, candidates):
    """Case-insensitive exact match, returns the canonical candidate or None."""
    wanted = (value or "").strip().lower()
    for name in candidates:
        if name.strip().lower() == wanted:
            return name
    return None


def plan_targets(cfg, folders, delimiter, existing_names):
    """Returns (to_create, reused, warnings) as sets of lowercase wire names."""
    existing = {name.lower() for name in existing_names}
    create, reuse, warnings = set(), set(), []

    if cfg.target_parent_folder:
        for folder in folders:
            path = mig.build_target_path(cfg, folder["path"], delimiter)
            current = ""
            for part in path.split(delimiter):
                current = part if not current else f"{current}{delimiter}{part}"
                wire = mig.imap_utf7_encode(current).lower()
                (reuse if wire in existing else create).add(wire)
    else:
        inbox = mig.imap_utf7_encode("INBOX").lower()
        for folder in folders:
            target = mig.TARGET_FOLDER_MAP.get(folder["displayName"], "INBOX")
            wire = mig.imap_utf7_encode(target).lower()
            if wire in existing:
                reuse.add(wire)
            else:
                warnings.append(f"{folder['path']}: target folder '{target}' is missing, "
                                f"messages go to INBOX")
                reuse.add(inbox)
    return create, reuse, warnings


def build_plan(cfg, mailbox, folders, delimiter, existing_names):
    create, reuse, warnings = plan_targets(cfg, folders, delimiter, existing_names)
    total = sum(folder.get("totalItemCount") or 0 for folder in folders)
    if cfg.migrate_limit:
        total = min(total, cfg.migrate_limit)

    parent_exists = False
    if cfg.target_parent_folder:
        parent_exists = (mig.imap_utf7_encode(cfg.target_parent_folder).lower()
                         in {name.lower() for name in existing_names})
        if not parent_exists:
            warnings.insert(0, f"target folder '{cfg.target_parent_folder}' will be created")

    return Plan(
        source_email=mailbox.source_email,
        target_email=mailbox.target_email,
        target_server=mailbox.target_server,
        folders=[folder["path"] for folder in folders],
        total_messages=total,
        limit=cfg.migrate_limit,
        parent_folder=cfg.target_parent_folder,
        parent_exists=parent_exists,
        create_count=len(create),
        reuse_count=len(reuse),
        attachments=cfg.migrate_attachments,
        email_format=cfg.email_format,
        imap_ssl_verify=cfg.imap_ssl_verify,
        warnings=warnings,
    )


def plan_rows(plan):
    if len(plan.folders) <= 5:
        folders = ", ".join(plan.folders)
    else:
        folders = ", ".join(plan.folders[:4]) + f", ... ({len(plan.folders)} total)"

    if plan.parent_folder:
        parent = plan.parent_folder + (" [existing]" if plan.parent_exists
                                       else " [will be created]")
    else:
        parent = "(none - classic folder mapping)"

    return [
        ("Source", f"{plan.source_email}  ({plan.total_messages} message(s), "
                   f"{len(plan.folders)} folder(s))"),
        ("Target", f"{plan.target_email} @ {plan.target_server}"),
        ("Folders", folders),
        ("Limit", f"newest {plan.limit} message(s)" if plan.limit else "all messages"),
        ("Attachments", "on" if plan.attachments else "off"),
        ("Body format", plan.email_format),
        ("Target parent", parent),
        ("Target folders", f"{plan.create_count} to create, {plan.reuse_count} reused"),
        ("IMAP TLS", "verified" if plan.imap_ssl_verify else "NOT verified"),
    ]


def write_env_values(path, values):
    """Atomically update KEY=value lines, preserving comments and order."""
    lines = []
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as handle:
            lines = handle.read().splitlines()

    remaining = dict(values)
    for index, line in enumerate(lines):
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        key = stripped.split("=", 1)[0].strip()
        if key in remaining:
            lines[index] = f"{key}={remaining.pop(key)}"

    for key, value in remaining.items():
        lines.append(f"{key}={value}")

    tmp_path = path + ".tmp"
    with open(tmp_path, "w", encoding="utf-8") as handle:
        handle.write("\n".join(lines) + "\n")
    os.chmod(tmp_path, 0o600)
    os.replace(tmp_path, path)
    return path


def append_csv_row(path, row):
    """Appends a mailbox row unless an equivalent row exists. Returns True if written."""
    row = tuple(str(value) for value in row)
    normalized = tuple(value.strip().lower() for value in row)
    rows = []
    if os.path.exists(path):
        with open(path, newline="", encoding="utf-8-sig") as handle:
            rows = [existing for existing in csv.reader(handle)
                    if any(cell.strip() for cell in existing)]

    for existing in rows[1:]:
        if len(existing) == len(row) and \
                tuple(cell.strip().lower() for cell in existing) == normalized:
            return False

    write_header = not rows
    with open(path, "a", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        if write_header:
            writer.writerow(mig.CSV_HEADER)
            os.chmod(path, 0o600)
        writer.writerow(row)
    return True


def verify_credentials(cfg):
    """Returns None when the app credentials work, otherwise an error string."""
    preflight_dir = os.path.join(mig.TOKENS_DIR, "_preflight")
    shutil.rmtree(preflight_dir, ignore_errors=True)
    try:
        account = mig.Account(
            (cfg.client_id, cfg.client_secret),
            tenant_id=cfg.tenant_id,
            token_backend=mig.FileSystemTokenBackend(token_path=preflight_dir),
            auth_flow_type="credentials",
        )
        if not account.is_authenticated and not account.authenticate():
            return "authentication failed"
        return None
    except Exception as exc:
        return str(exc) or exc.__class__.__name__
    finally:
        shutil.rmtree(preflight_dir, ignore_errors=True)


# Wizard ---------------------------------------------------------------------

class Wizard:
    def __init__(self, cfg, prompter, renderer, env_path=mig.DEFAULT_ENV):
        self.cfg = cfg
        self.prompter = prompter
        self.renderer = renderer
        self.env_path = env_path
        self.mailbox = MailboxChoice("")
        self.folders = []
        self.selected = []
        self.delimiter = "/"
        self.target_names = []
        self.mail = None
        self.access_token = None
        self.save_row = False

    # -- steps -------------------------------------------------------------

    def step_mailbox(self):
        self.renderer.step(1, 8, "Source mailbox")
        rows = []
        if os.path.exists(self.cfg.csv_file):
            try:
                rows = mig.read_mailboxes_from_csv(self.cfg.csv_file)
            except mig.MigrationError as exc:
                self.renderer.warn(str(exc))

        if rows:
            options = [f"{source} -> {target} @ {server}"
                       for source, server, target, _ in rows]
            options.append("Enter a different mailbox...")
            index = self.prompter.select("Which mailbox do you want to migrate?",
                                         options, default=0)
            if index < len(rows):
                source, server, target, password = rows[index]
                self.mailbox = MailboxChoice(source, server, target, password,
                                             from_csv=True)
                self.renderer.info(f"Using {source} -> {target} @ {server}")

        if not self.mailbox.source_email:
            source_email = self.prompter.text("Source O365 mailbox address",
                                              validate=validate_email)
            self.mailbox = MailboxChoice(source_email)

    def step_credentials(self):
        self.renderer.step(2, 8, "Microsoft Graph credentials")
        self.cfg = credentials_flow(self.cfg, self.prompter, self.renderer, self.env_path)

    def step_stats(self):
        self.renderer.step(3, 8, "Account statistics")
        try:
            self.access_token, _ = mig.authenticate_account(
                self.cfg, self.mailbox.source_email, force_reauth=False)
            self.folders = mig.get_mail_folders(
                self.cfg, self.access_token, self.mailbox.source_email)
        except Exception as exc:
            raise InteractiveError(
                f"cannot read mailbox {self.mailbox.source_email}: {exc}") from exc

        if not self.folders:
            raise InteractiveError("the mailbox has no folders")

        self.renderer.stats(self.folders)
        if not any(folder.get("totalItemCount") for folder in self.folders):
            self.renderer.warn("the mailbox appears to be empty")

    def step_scope(self):
        self.renderer.step(4, 8, "Scope")
        paths = [folder["path"] for folder in self.folders]
        prefilled = {name.strip().lower() for name in self.cfg.migrate_folders}
        preselected = tuple(index for index, folder in enumerate(self.folders)
                            if folder["path"].lower() in prefilled)

        choice = self.prompter.select(
            f"Which folders? ({len(self.folders)} available)",
            [f"All folders ({len(self.folders)})", "Choose folders..."],
            default=1 if preselected else 0)

        if choice == 0:
            self.selected = list(self.folders)
        else:
            indexes = self.prompter.multi_select(
                "Select source folders (a parent includes its subfolders)",
                paths, preselected=preselected)
            if not indexes:
                raise InteractiveError("no folders selected")
            chosen = {self.folders[index]["path"] for index in indexes}
            self.selected = [
                folder for folder in self.folders
                if any(folder["path"] == path or folder["path"].startswith(path + "/")
                       for path in chosen)
            ]

        limit = self.cfg.migrate_limit
        if self.prompter.confirm("Only migrate the newest N messages per selection?",
                                 default=bool(limit)):
            answer = self.prompter.text("Number of messages",
                                        default=str(limit or ""),
                                        validate=validate_positive_int)
            limit = int(answer)
        else:
            limit = 0

        self.cfg = replace(
            self.cfg,
            migrate_folders=[folder["path"] for folder in self.selected],
            migrate_limit=limit,
        )

    def step_imap(self):
        self.renderer.step(5, 8, "Target IMAP")
        server = self.mailbox.target_server
        target = self.mailbox.target_email
        password = self.mailbox.password

        if self.mailbox.from_csv and target and server:
            if self.prompter.confirm(f"Target is {target} @ {server}. Change it?",
                                     default=False):
                server = target = password = ""
        if not server:
            server = self.prompter.text("IMAP server", default=server,
                                        validate=lambda v: v or _raise("enter a server"))
        if not target:
            target = self.prompter.text("Target mailbox address", validate=validate_email)
        if not password:
            password = self.prompter.secret(f"IMAP password for {target}")

        while True:
            mail = mig.connect_to_target_imap(self.cfg, server, target, password)
            if mail is not None:
                break
            if not self.prompter.confirm("Login failed. Try different credentials?",
                                         default=True):
                raise InteractiveError("IMAP login failed")
            server = self.prompter.text("IMAP server", default=server,
                                        validate=lambda v: v or _raise("enter a server"))
            target = self.prompter.text("Target mailbox address", default=target,
                                        validate=validate_email)
            password = self.prompter.secret("IMAP password")

        self.mail = mail
        self.delimiter, self.target_names = mig.get_imap_hierarchy(mail)
        self.renderer.info(f"Connected: {len(self.target_names)} target folder(s), "
                           f"delimiter {self.delimiter!r}.")
        if not self.mailbox.from_csv:
            self.save_row = self.prompter.confirm(
                "Save this mailbox to details.csv after a successful migration?",
                default=False)
        self.mailbox = replace(self.mailbox, target_server=server,
                               target_email=target, password=password)

    def step_parent(self):
        self.renderer.step(6, 8, "Target parent folder")
        if not self.prompter.confirm(
                "Mirror the O365 folder tree under a target folder?",
                default=bool(self.cfg.target_parent_folder)):
            self.cfg = replace(self.cfg, target_parent_folder="")
            return

        roots = root_folder_candidates(self.target_names, self.delimiter)
        if len(roots) < 10:
            options = [f"{name}{' (special)' if is_special_root(name) else ''}"
                       for name in roots]
            options.append("Type a new folder name...")
            index = self.prompter.select("Select a root folder on the target",
                                         options, default=0)
            if index < len(roots):
                self.cfg = replace(self.cfg, target_parent_folder=roots[index])
                self.renderer.info(f"Parent folder: {roots[index]} (existing)")
                return

        value = self.prompter.text(
            "Target parent folder",
            default=self.cfg.target_parent_folder,
            validate=lambda v: validate_parent_name(v, self.delimiter),
            completer=roots,
        )
        matched = match_folder(value, roots)
        chosen = matched or value.strip()
        self.cfg = replace(self.cfg, target_parent_folder=chosen)
        self.renderer.info(f"Parent folder: {chosen} "
                           f"({'existing' if matched else 'will be created'})")

    # -- validation and plan ----------------------------------------------

    def _sample_check(self):
        try:
            candidates = [folder for folder in self.selected
                          if folder.get("totalItemCount")] or self.selected
            if not candidates:
                return Check("Sample message", False, "no folders selected")
            folder = candidates[0]
            page = mig.graph_get(
                f"{mig._user_url(self.mailbox.source_email)}"
                f"/mailFolders/{folder['id']}/messages",
                self.access_token, params={"$top": 1})
            messages = page.get("value") or []
            if not messages:
                return Check("Sample message", True, "selection is empty")

            message = messages[0]
            message["folderName"] = folder["displayName"]
            message["folderPath"] = folder["path"]
            if self.cfg.migrate_attachments:
                message["attachments"] = mig.fetch_attachments(
                    self.access_token, self.mailbox.source_email, message["id"])
            rfc822 = mig.convert_to_rfc822(self.cfg, message)
            if not rfc822:
                return Check("Sample message", False, "message conversion failed")
            return Check("Sample message", True,
                         f"{folder['path']}, {len(rfc822)} bytes")
        except Exception as exc:
            return Check("Sample message", False, str(exc) or exc.__class__.__name__)

    def validate_and_plan(self):
        checks = [
            Check("Graph app credentials", bool(self.cfg.client_id), self.cfg.tenant_id),
            Check("Mailbox access", bool(self.folders),
                  f"{len(self.folders)} folder(s) readable"),
            self._sample_check(),
            Check("Target IMAP login", self.mail is not None,
                  f"{self.mailbox.target_server} "
                  f"({'TLS verified' if self.cfg.imap_ssl_verify else 'TLS NOT verified'})"),
        ]
        plan = build_plan(self.cfg, self.mailbox, self.selected,
                          self.delimiter, self.target_names)
        checks.append(Check("Folder plan", True,
                            f"{plan.create_count} to create, {plan.reuse_count} reused"))
        return checks, plan

    # -- execution ---------------------------------------------------------

    def close(self):
        if self.mail is not None:
            try:
                self.mail.logout()
            except Exception:
                pass
            self.mail = None

    def rerun_command(self):
        parts = ["python", "o365_imap_migration.py"]
        if self.cfg.csv_file != mig.DEFAULT_CSV:
            parts += ["--csv", self.cfg.csv_file]
        if self.cfg.migrate_folders:
            parts += ["--folders", ",".join(self.cfg.migrate_folders)]
        if self.cfg.migrate_limit:
            parts += ["--limit", str(self.cfg.migrate_limit)]
        if self.cfg.target_parent_folder:
            parts += ["--target-parent-folder", self.cfg.target_parent_folder]
        return shlex.join(parts)

    def execute(self):
        self.renderer.step(8, 8, "Migration")
        self.close()

        board = mig.ProgressBoard()
        previous_board = mig._active_board
        mig._active_board = board
        started = time.time()
        try:
            migrated, failed = mig.migrate_mailbox(
                self.cfg,
                (self.mailbox.source_email, self.mailbox.target_server,
                 self.mailbox.target_email, self.mailbox.password),
                board,
            )
        finally:
            mig._active_board = previous_board

        elapsed = time.time() - started
        self.renderer.info(f"Migration finished in {elapsed:.1f}s: "
                           f"{migrated} migrated, {failed} failed.")

        if self.save_row and not self.mailbox.from_csv:
            row = (self.mailbox.source_email, self.mailbox.target_server,
                   self.mailbox.target_email, self.mailbox.password)
            if append_csv_row(self.cfg.csv_file, row):
                self.renderer.info(f"Mailbox appended to {self.cfg.csv_file}.")
            else:
                self.renderer.info(f"Mailbox already present in {self.cfg.csv_file}.")

        self.renderer.info(f"Rerun non-interactively: {self.rerun_command()}")
        return 0 if failed == 0 else 1

    # -- orchestration -----------------------------------------------------

    def run(self):
        self.step_mailbox()
        self.step_credentials()
        self.step_stats()
        self.step_imap()
        self.step_parent()

        while True:
            self.step_scope()
            checks, plan = self.validate_and_plan()
            self.renderer.plan(checks, plan)
            if all(check.ok for check in checks):
                break
            if not self.prompter.confirm(
                    "Validation failed. Go back and adjust the selection?", default=True):
                self.close()
                self.renderer.info("Aborted, nothing written.")
                return 2

        if not self.prompter.confirm("Start the migration?", default=False):
            self.close()
            self.renderer.info("Aborted, nothing written.")
            return 2

        return self.execute()


def _raise(message):
    raise ValueError(message)


def credentials_flow(cfg, prompter, renderer, env_path):
    """Verify ``SOURCE_*`` credentials, prompting (and optionally persisting) new ones."""
    attempts = 0
    while attempts < 3:
        if cfg.client_id and cfg.client_secret and cfg.tenant_id:
            error = verify_credentials(cfg)
            if error is None:
                renderer.info("Microsoft Graph credentials verified.")
                return cfg
            attempts += 1
            renderer.warn(f"Graph credentials rejected: {error}")
            if attempts >= 3:
                break
        else:
            renderer.info("Graph app credentials are missing from .env.")

        client_id = prompter.text("Client ID", default=cfg.client_id)
        client_secret = prompter.secret("Client secret")
        tenant_id = prompter.text("Tenant ID", default=cfg.tenant_id or "common")
        cfg = replace(cfg, client_id=client_id, client_secret=client_secret,
                      tenant_id=tenant_id)

        error = verify_credentials(cfg)
        if error is None:
            renderer.info("Microsoft Graph credentials verified.")
            if prompter.confirm("Save these credentials to .env?", default=False):
                write_env_values(env_path, {
                    "SOURCE_CLIENT_ID": client_id,
                    "SOURCE_CLIENT_SECRET": client_secret,
                    "SOURCE_TENANT_ID": tenant_id,
                })
                renderer.info(f"Credentials saved to {env_path}.")
            return cfg

        attempts += 1
        renderer.warn(f"Graph credentials rejected: {error}")

    raise InteractiveError("giving up after 3 failed credential attempts")


# Entry point ----------------------------------------------------------------

def run_interactive(args, prompter=None, renderer=None):
    """Runs the wizard. Returns the process exit code."""
    if prompter is None or renderer is None:
        if not (sys.stdin.isatty() and sys.stdout.isatty()):
            print("Interactive mode needs a terminal. "
                  "Use the non-interactive flags instead.", file=sys.stderr)
            return 2
        try:
            renderer = RichRenderer()
            prompter = PromptToolkitPrompter()
        except InteractiveError as exc:
            print(f"Error: {exc}", file=sys.stderr)
            return 2

    try:
        cfg = mig.load_config(args.env_file, require_credentials=False)
    except mig.MigrationError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2

    if args.csv:
        cfg = replace(cfg, csv_file=args.csv)
    if getattr(args, "target_parent_folder", None) is not None:
        cfg = replace(cfg, target_parent_folder=mig.normalize_parent_folder(
            args.target_parent_folder))
    if getattr(args, "folders", None) is not None:
        cfg = replace(cfg, migrate_folders=mig.parse_folder_list(args.folders))
    if getattr(args, "limit", None) is not None:
        if args.limit < 0:
            print("Error: --limit must be 0 or a positive integer", file=sys.stderr)
            return 2
        cfg = replace(cfg, migrate_limit=args.limit)

    renderer.info("Interactive migration - one mailbox. Ctrl-C aborts.")
    if getattr(args, "dry_run", False):
        renderer.info("Note: interactive mode always validates before migrating, "
                      "so --dry-run is ignored here.")

    wizard = Wizard(cfg, prompter, renderer, args.env_file or mig.DEFAULT_ENV)
    try:
        return wizard.run()
    except InteractiveError as exc:
        wizard.close()
        renderer.error(str(exc))
        return 2
    except (KeyboardInterrupt, EOFError):
        wizard.close()
        renderer.info("\nInterrupted, nothing written.")
        return 130
