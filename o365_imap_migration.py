#!/usr/bin/env python3
"""Migrate Microsoft 365 mailboxes to an IMAP server.

Configuration is read from .env (see .env.example); the mailboxes to migrate
are read from details.csv (see details.csv.example).

    python o365_imap_migration.py --dry-run   # verify credentials/access, write nothing
    python o365_imap_migration.py             # migrate
"""

from __future__ import annotations

import argparse
import base64
import csv
import hashlib
import html
import imaplib
import json
import os
import re
import shutil
import ssl
import sys
import threading
import time
from dataclasses import dataclass, field
from datetime import datetime, timezone
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import format_datetime, parsedate_to_datetime
from typing import Optional
from urllib.parse import quote

try:
    import requests
    from dotenv import load_dotenv
    from O365 import Account, FileSystemTokenBackend
except ImportError as exc:  # pragma: no cover - dependency check
    sys.exit(
        f"Missing dependency: {exc.name}. "
        "Install the requirements with: python -m pip install -r requirements.txt"
    )

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKENS_DIR = "tokens"
SESSIONS_DIR = "sessions"
DEFAULT_ENV = ".env"
DEFAULT_CSV = "details.csv"
CSV_HEADER = ["source_email", "target_server", "target_email", "target_password"]
REQUIRED_ENV = ["SOURCE_CLIENT_ID", "SOURCE_CLIENT_SECRET", "SOURCE_TENANT_ID"]
TRUE_VALUES = {"1", "true", "yes", "y", "on"}
RETRY_STATUS_CODES = {429, 500, 502, 503, 504}

# Fallback mapping for the default folders; anything unmapped goes to INBOX.
TARGET_FOLDER_MAP = {
    "Inbox": "INBOX",
    "Drafts": "Drafts",
    "Sent Items": "Sent",
    "Deleted Items": "Deleted Items",
    "Junk Email": "Spam",
}


class MigrationError(Exception):
    """Fatal configuration or authentication error."""


class GraphError(Exception):
    """A Microsoft Graph request failed."""


# Printing and progress ------------------------------------------------------

print_lock = threading.Lock()
_active_board: Optional["ProgressBoard"] = None


def safe_print(message, end="\n"):
    """Print from any thread; routed through the progress board when active."""
    board = _active_board
    if board is not None:
        board.notify(message)
        return
    with print_lock:
        print(message, end=end, flush=True)


class ProgressBoard:
    """Per-mailbox migration progress.

    On a terminal the board redraws a block of one line per mailbox in place.
    When the output is redirected it emits one line per mailbox every few
    seconds (and on completion), which keeps CI logs readable.
    """

    def __init__(self, stream=None, tty=None, interval=2.0):
        self.stream = stream or sys.stdout
        self.tty = self.stream.isatty() if tty is None else tty
        self.interval = interval
        self.lock = threading.RLock()
        self.rows = {}
        self.order = []
        self.drawn = 0

    # -- public API ---------------------------------------------------------

    def register(self, label, total):
        """Adds a mailbox row and returns its (unique) key."""
        with self.lock:
            key = label
            suffix = 2
            while key in self.rows:
                key = f"{label} #{suffix}"
                suffix += 1
            self.rows[key] = {
                "total": max(int(total), 0),
                "done": 0,
                "migrated": 0,
                "failed": 0,
                "started": time.monotonic(),
                "last_emit": 0.0,
                "folder": "",
                "status": "running",
            }
            self.order.append(key)
            self._render(key, force=True)
            return key

    def set_folder(self, key, folder):
        with self.lock:
            if key in self.rows:
                self.rows[key]["folder"] = folder
                self._render(key)

    def advance(self, key, ok=True):
        with self.lock:
            row = self.rows.get(key)
            if row is None:
                return
            row["done"] += 1
            if ok:
                row["migrated"] += 1
            else:
                row["failed"] += 1
            self._render(key)

    def finish(self, key, status="done"):
        with self.lock:
            row = self.rows.get(key)
            if row is None:
                return
            row["status"] = status
            if not self.tty:
                self._render(key, force=True)
                return
            self._draw_all()
            # Commit the block to the scrollback once every mailbox is done.
            if all(item["status"] != "running" for item in self.rows.values()):
                self._clear()
                for order_key in self.order:
                    self.stream.write(self._format(order_key) + "\n")
                self.drawn = 0
                self.stream.flush()

    def log(self, key, message):
        self.notify(f"[{key}] {message}")

    def notify(self, message):
        """Prints a message without corrupting the redrawn block."""
        with self.lock:
            if self.tty:
                self._clear()
            self.stream.write(str(message) + "\n")
            if self.tty:
                self._draw_all()
            else:
                self.stream.flush()

    # -- rendering ----------------------------------------------------------

    def _format(self, key):
        row = self.rows[key]
        total = row["total"]
        done = row["done"]
        percent = (100.0 * done / total) if total else 100.0
        elapsed = max(time.monotonic() - row["started"], 1e-9)
        rate = done / elapsed

        if done and done < total:
            seconds = int((total - done) / rate) if rate else 0
            eta = f"ETA {seconds // 60}d{seconds % 60:02d}s"
        elif done >= total:
            eta = "ETA done"
        else:
            eta = "ETA --"

        folder = row["folder"] or "-"
        if len(folder) > 32:
            folder = "..." + folder[-29:]

        return (f"[{key}] {done}/{total} {percent:5.1f}% | {row['migrated']} ok, "
                f"{row['failed']} failed | {rate:5.1f} msg/s | {eta} | {folder} | "
                f"{row['status']}")

    def _render(self, key, force=False):
        if self.tty:
            self._draw_all()
            return

        row = self.rows[key]
        now = time.monotonic()
        if not force and row["done"] < row["total"] and (now - row["last_emit"]) < self.interval:
            return
        row["last_emit"] = now
        self.stream.write(self._format(key) + "\n")
        self.stream.flush()

    def _clear(self):
        if not self.drawn:
            return
        self.stream.write(f"\x1b[{self.drawn}A")
        for _ in range(self.drawn):
            self.stream.write("\x1b[2K\n")
        self.stream.write(f"\x1b[{self.drawn}A")
        self.drawn = 0

    def _draw_all(self):
        self._clear()
        for key in self.order:
            self.stream.write("\x1b[2K" + self._format(key) + "\n")
        self.drawn = len(self.order)
        self.stream.flush()


# Configuration --------------------------------------------------------------

@dataclass
class Config:
    client_id: str
    client_secret: str
    tenant_id: str
    migrate_attachments: bool = True
    email_format: str = "html"
    imap_ssl_verify: bool = True
    graph_page_size: int = 100
    csv_file: str = DEFAULT_CSV
    target_parent_folder: str = ""
    migrate_folders: list[str] = field(default_factory=list)
    migrate_limit: int = 0  # 0 = everything
    migrate_after: str = ""   # receivedDateTime >= (UTC ISO-8601)
    migrate_before: str = ""  # receivedDateTime <  (UTC ISO-8601)
    restartable: bool = False
    reset_session: bool = False
    session_dir: str = SESSIONS_DIR


def env_bool(name, default=False):
    raw = os.getenv(name)
    if raw is None or not raw.strip():
        return default
    return raw.strip().lower() in TRUE_VALUES


def env_int(name, default, low, high):
    raw = os.getenv(name)
    if raw is None or not raw.strip():
        return default
    try:
        value = int(raw.strip())
    except ValueError:
        raise MigrationError(f"{name} must be an integer, got {raw!r}") from None
    if not low <= value <= high:
        raise MigrationError(f"{name} must be between {low} and {high}, got {value}")
    return value


def parse_folder_list(value):
    """Splits a comma-separated folder list, dropping empty entries."""
    return [item.strip() for item in (value or "").split(",") if item.strip()]


def normalize_parent_folder(value):
    """Normalizes a target parent folder to 'A/B' form (no leading/trailing slash)."""
    return "/".join(part.strip() for part in (value or "").split("/") if part.strip())


def normalize_date_bound(value):
    """Normalizes YYYY-MM-DD or ISO-8601 to a UTC ``YYYY-MM-DDTHH:MM:SSZ`` string."""
    text = (value or "").strip()
    if not text:
        return ""
    try:
        if len(text) == 10:
            parsed = datetime.strptime(text, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        else:
            parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        raise MigrationError(
            f"invalid date {value!r}; use YYYY-MM-DD or an ISO-8601 timestamp") from None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def build_received_filter(cfg):
    """Graph ``$filter`` expression for the configured received-date range."""
    after = normalize_date_bound(cfg.migrate_after)
    before = normalize_date_bound(cfg.migrate_before)
    clauses = []
    if after:
        clauses.append(f"receivedDateTime ge {after}")
    if before:
        clauses.append(f"receivedDateTime lt {before}")
    return " and ".join(clauses)


def load_config(env_file=None, require_credentials=True):
    """Load settings from a .env file (and the process environment).

    With ``require_credentials=False`` missing ``SOURCE_*`` values are allowed and
    returned empty; interactive mode uses that to prompt for them.
    """
    path = env_file or DEFAULT_ENV
    if env_file and not os.path.exists(path):
        raise MigrationError(f"env file not found: {path}")
    if os.path.exists(path):
        # An explicitly requested env file wins over the process environment.
        load_dotenv(path, override=bool(env_file))

    if require_credentials:
        missing = [name for name in REQUIRED_ENV if not (os.getenv(name) or "").strip()]
        if missing:
            hint = "" if os.path.exists(path) else \
                f" (copy {DEFAULT_ENV}.example to {DEFAULT_ENV} first)"
            raise MigrationError("missing required setting(s): " + ", ".join(missing) + hint)

    email_format = (os.getenv("EMAIL_FORMAT") or "html").strip().lower()
    if email_format not in ("html", "plain"):
        raise MigrationError(f"EMAIL_FORMAT must be 'html' or 'plain', got {email_format!r}")

    migrate_after = normalize_date_bound(os.getenv("MIGRATE_AFTER", ""))
    migrate_before = normalize_date_bound(os.getenv("MIGRATE_BEFORE", ""))
    if migrate_after and migrate_before and migrate_after >= migrate_before:
        raise MigrationError("MIGRATE_AFTER must be earlier than MIGRATE_BEFORE")

    return Config(
        client_id=(os.getenv("SOURCE_CLIENT_ID") or "").strip(),
        client_secret=(os.getenv("SOURCE_CLIENT_SECRET") or "").strip(),
        tenant_id=(os.getenv("SOURCE_TENANT_ID") or "").strip(),
        migrate_attachments=env_bool("MIGRATE_ATTACHMENTS", True),
        email_format=email_format,
        imap_ssl_verify=env_bool("IMAP_SSL_VERIFY", True),
        graph_page_size=env_int("GRAPH_PAGE_SIZE", 100, 1, 1000),
        csv_file=(os.getenv("CSV_FILE") or DEFAULT_CSV).strip(),
        target_parent_folder=normalize_parent_folder(os.getenv("TARGET_PARENT_FOLDER", "")),
        migrate_folders=parse_folder_list(os.getenv("MIGRATE_FOLDERS", "")),
        migrate_limit=env_int("MIGRATE_LIMIT", 0, 0, 1_000_000_000),
        migrate_after=migrate_after,
        migrate_before=migrate_before,
        restartable=env_bool("RESTARTABLE", False),
        reset_session=env_bool("RESET_SESSION", False),
        session_dir=(os.getenv("SESSION_DIR") or SESSIONS_DIR).strip(),
    )


def read_mailboxes_from_csv(filename):
    """Reads mailboxes from a CSV file (source_email,target_server,target_email,target_password)."""
    if not os.path.exists(filename):
        raise MigrationError(
            f"{filename} not found. Copy details.csv.example to {filename} and fill it in."
        )

    mailboxes = []
    seen = set()
    with open(filename, newline="", encoding="utf-8-sig") as csvfile:
        reader = csv.reader(csvfile)
        try:
            header = [cell.strip().lower() for cell in next(reader)]
        except StopIteration:
            raise MigrationError(f"{filename} is empty.") from None
        if header != CSV_HEADER:
            safe_print(f"Warning: unexpected header in {filename}: {header}. "
                       f"Expected: {CSV_HEADER}")

        for line_number, row in enumerate(reader, start=2):
            row = [cell.strip() for cell in row]
            if not any(row):
                continue
            if len(row) != len(CSV_HEADER):
                safe_print(f"Warning: skipping {filename}:{line_number}: expected "
                           f"{len(CSV_HEADER)} columns, got {len(row)}.")
                continue
            source_email, target_server, target_email, target_password = row
            if not all(row):
                safe_print(f"Warning: skipping {filename}:{line_number}: empty value in row.")
                continue
            key = (source_email.lower(), target_server.lower(), target_email.lower())
            if key in seen:
                safe_print(f"Warning: skipping {filename}:{line_number}: duplicate entry for "
                           f"{source_email} -> {target_email}.")
                continue
            seen.add(key)
            mailboxes.append((source_email, target_server, target_email, target_password))
    return mailboxes


# Restartable sessions -------------------------------------------------------

def utcnow_iso():
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def session_path(cfg, mailbox):
    """Deterministic per-mailbox session file path."""
    source, server, target, _ = mailbox
    slug = re.sub(r"[^A-Za-z0-9._@-]+", "_", f"{source}__{target}__{server}")[:80]
    key = "|".join(value.strip().lower() for value in (source, target, server))
    digest = hashlib.sha1(key.encode("utf-8")).hexdigest()[:8]
    return os.path.join(cfg.session_dir, f"{slug}__{digest}.jsonl")


class MigrationSession:
    """Append-only record of message IDs already appended to the target.

    One JSON object per line; a torn last line (crash during write) is ignored.
    The first line is a header that identifies the source/target/parent folder so a
    session from a different migration is never mixed in by accident.
    """

    VERSION = 1

    def __init__(self, path, source, target, server, parent_folder):
        self.path = path
        self.source = source
        self.target = target
        self.server = server
        self.parent_folder = parent_folder
        self.ids = set()
        self.written = 0
        self._handle = None
        self._since_sync = 0

    @classmethod
    def load(cls, cfg, mailbox):
        source, server, target, _ = mailbox
        session = cls(session_path(cfg, mailbox), source, target, server,
                      cfg.target_parent_folder)
        if not os.path.exists(session.path):
            return session

        header = None
        with open(session.path, encoding="utf-8", errors="replace") as handle:
            for line in handle:
                line = line.strip()
                if not line:
                    continue
                try:
                    entry = json.loads(line)
                except ValueError:
                    continue  # tolerate a partially written trailing line
                if header is None:
                    if entry.get("type") == "header":
                        header = entry
                    continue
                if entry.get("id"):
                    session.ids.add(str(entry["id"]))

        if header is None:
            safe_print(f"Warning: ignoring session file without header: {session.path}")
            session.ids.clear()
            return session

        expected = (source, target, server, cfg.target_parent_folder)
        actual = (header.get("source", ""), header.get("target", ""),
                  header.get("server", ""), header.get("parent_folder", ""))
        if (tuple(value.strip().lower() for value in actual)
                != tuple(value.strip().lower() for value in expected)):
            raise MigrationError(
                f"session file {session.path} belongs to a different migration "
                f"({actual[0]} -> {actual[1]} @ {actual[2]}, parent {actual[3]!r}); "
                f"use --reset-session to start over")
        return session

    def reset(self):
        """Deletes the session file (used by --reset-session)."""
        self._close()
        if os.path.exists(self.path):
            os.remove(self.path)
        self.ids.clear()
        self.written = 0
        return self

    def contains(self, message_id):
        return str(message_id) in self.ids

    def record(self, message_id, folder=""):
        message_id = str(message_id)
        if message_id in self.ids:
            return
        self._ensure_open()
        entry = {"id": message_id, "folder": folder, "ts": utcnow_iso()}
        self._handle.write(json.dumps(entry, separators=(",", ":")) + "\n")
        self._handle.flush()
        self.ids.add(message_id)
        self.written += 1
        self._since_sync += 1
        if self._since_sync >= 32:
            os.fsync(self._handle.fileno())
            self._since_sync = 0

    def _ensure_open(self):
        if self._handle is not None:
            return
        directory = os.path.dirname(self.path)
        if directory:
            os.makedirs(directory, exist_ok=True)
        new_file = not os.path.exists(self.path)
        self._handle = open(self.path, "a", encoding="utf-8")
        if new_file:
            header = {
                "type": "header", "version": self.VERSION,
                "source": self.source, "target": self.target,
                "server": self.server, "parent_folder": self.parent_folder,
                "created": utcnow_iso(),
            }
            self._handle.write(json.dumps(header, separators=(",", ":")) + "\n")
            self._handle.flush()
            try:
                os.chmod(self.path, 0o600)
            except OSError:
                pass

    def _close(self):
        if self._handle is not None:
            try:
                self._handle.flush()
                os.fsync(self._handle.fileno())
            except OSError:
                pass
            self._handle.close()
            self._handle = None

    def close(self):
        self._close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()


# Microsoft Graph ------------------------------------------------------------

def _user_url(email_address):
    return f"{GRAPH_BASE}/users/{quote(email_address, safe='@')}"


def _retry_delay(response, attempt):
    header = response.headers.get("Retry-After") if response is not None else None
    if header:
        try:
            return max(float(header), 1.0)
        except ValueError:
            pass
    return min(2 ** attempt, 30)


def graph_get(url, access_token, params=None, timeout=60, max_attempts=5):
    """Single Graph GET with retry/backoff for throttling and transient errors."""
    headers = {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}
    last_error = "unknown error"

    for attempt in range(1, max_attempts + 1):
        try:
            response = requests.get(url, headers=headers, params=params, timeout=timeout)
        except requests.RequestException as e:
            last_error = f"{type(e).__name__}: {e}"
            if attempt == max_attempts:
                break
            delay = min(2 ** attempt, 30)
            safe_print(f"Graph request failed ({last_error}); retrying in {delay}s "
                       f"({attempt}/{max_attempts})")
            time.sleep(delay)
            continue

        if response.status_code == 200:
            return response.json()

        last_error = f"HTTP {response.status_code}: {response.text[:300]}"
        if response.status_code not in RETRY_STATUS_CODES:
            raise GraphError(f"GET {url} failed: {last_error}")
        if attempt == max_attempts:
            break
        delay = _retry_delay(response, attempt)
        safe_print(f"Graph returned {response.status_code}; retrying in {delay:.0f}s "
                   f"({attempt}/{max_attempts})")
        time.sleep(delay)

    raise GraphError(f"GET {url} failed after {max_attempts} attempts: {last_error}")


def graph_get_all(url, access_token, params=None, timeout=60, limit=None,
                  on_page=None, skip_ids=None):
    """GET every page of a Graph collection by following @odata.nextLink.

    Graph returns at most 10 items per page by default, so ignoring
    @odata.nextLink silently drops everything after the first page.

    ``skip_ids`` removes already-migrated message IDs *before* the limit is
    applied, so a resumed run with ``--limit N`` fetches N *new* messages.
    """
    items = []
    done = False
    next_url, next_params = url, params
    while next_url and not done:
        page = graph_get(next_url, access_token, params=next_params, timeout=timeout)
        for item in page.get("value", []):
            if skip_ids and str(item.get("id")) in skip_ids:
                continue
            items.append(item)
            if limit is not None and len(items) >= limit:
                done = True
                break
        if on_page is not None:
            on_page(len(items))
        if done:
            break
        next_url = page.get("@odata.nextLink")
        # The nextLink already carries the query string of the original request.
        next_params = None
    if limit is not None:
        items = items[:limit]
    return items


def get_mail_folders(cfg, access_token, user_email):
    """Returns top-level and nested mail folders for a mailbox."""
    base = f"{_user_url(user_email)}/mailFolders"
    try:
        roots = graph_get_all(base, access_token, params={"$top": 100})
    except GraphError as e:
        raise GraphError(f"could not list mail folders: {e}") from e

    folders = []
    stack = [(folder, folder.get("displayName") or "Unknown") for folder in roots]
    while stack:
        folder, path = stack.pop()
        folders.append({
            "id": folder["id"],
            "displayName": folder.get("displayName") or "Unknown",
            "path": path,
            "totalItemCount": folder.get("totalItemCount") or 0,
        })
        try:
            children = graph_get_all(f"{base}/{folder['id']}/childFolders", access_token,
                                     params={"$top": 100})
        except GraphError as e:
            safe_print(f"Warning: could not list child folders of '{path}': {e}")
            continue
        stack.extend((child, f"{path}/{child.get('displayName') or 'Unknown'}")
                     for child in children)
    return folders


def select_folders(folders, requested):
    """Filters folders by path/display name; a match also selects its subtree."""
    wanted = list(dict.fromkeys(
        name.strip().strip("/").lower() for name in (requested or []) if name.strip()))
    if not wanted:
        return folders, []
    wanted_set = set(wanted)

    selected = []
    matched = set()
    for folder in folders:
        path = folder["path"].lower()
        parts = path.split("/")
        candidates = {path, folder["displayName"].lower()}
        for index in range(1, len(parts) + 1):
            candidates.add("/".join(parts[:index]))
        hit = candidates & wanted_set
        if hit:
            selected.append(folder)
            matched.update(hit)

    missing = [name for name in wanted if name not in matched]
    return selected, missing


def fetch_attachments(access_token, email_address, message_id):
    """Fetch attachments for a given message."""
    url = (f"{_user_url(email_address)}/messages/{quote(message_id, safe='')}/attachments")
    return graph_get_all(url, access_token)


def fetch_all_emails(cfg, source_email, access_token, skip_ids=None):
    """Fetch the selected emails from the source email account.

    ``skip_ids`` holds message IDs already recorded in a restartable session.
    """
    folders = get_mail_folders(cfg, access_token, source_email)
    selected, missing = select_folders(folders, cfg.migrate_folders)
    for name in missing:
        safe_print(f"[{source_email}] warning: source folder not found: {name!r}")
    if not selected:
        safe_print(f"[{source_email}] no source folders matched {cfg.migrate_folders!r}.")
        return []

    remaining = cfg.migrate_limit or None
    all_emails = []

    for folder in selected:
        if remaining is not None and remaining <= 0:
            break

        top = cfg.graph_page_size if remaining is None else min(cfg.graph_page_size, remaining)
        params = {"$top": top}
        filter_expression = build_received_filter(cfg)
        if filter_expression:
            params["$filter"] = filter_expression
        last_report = [0.0]

        def on_page(count, path=folder["path"]):
            now = time.monotonic()
            if now - last_report[0] >= 2.0:
                last_report[0] = now
                safe_print(f"[{source_email}] {path}: {count} message(s) fetched")

        try:
            messages = graph_get_all(
                f"{_user_url(source_email)}/mailFolders/{folder['id']}/messages",
                access_token,
                params=params,
                limit=remaining,
                on_page=on_page,
                skip_ids=skip_ids,
            )
        except GraphError as e:
            safe_print(f"[{source_email}] failed to fetch emails from folder "
                       f"'{folder['path']}': {e}")
            continue

        for message in messages:
            message["folderName"] = folder["displayName"]
            message["folderPath"] = folder["path"]
            if cfg.migrate_attachments:
                try:
                    message["attachments"] = fetch_attachments(
                        access_token, source_email, message["id"])
                except GraphError as e:
                    safe_print(f"[{source_email}] failed to fetch attachments for message "
                               f"{message.get('id')}: {e}")
                    message["attachments"] = []

        all_emails.extend(messages)
        safe_print(f"[{source_email}] {folder['path']}: {len(messages)} message(s) fetched")

        if remaining is not None:
            remaining -= len(messages)

    return all_emails


# Message conversion ---------------------------------------------------------

def extract_email_body(msg, email_format="html"):
    """Extract the body; strip markup only when the target expects plain text."""
    body = (msg.get("body") or {}).get("content") or ""
    if not body:
        return "(No Body)"

    if email_format == "plain":
        body = re.sub(r"<[^>]+>", "", body)
        body = re.sub(r"\s+", " ", body)
        body = html.unescape(body)

    return body.strip() or "(No Body)"


def _rfc2822_date(value):
    """Graph returns ISO 8601 timestamps; email headers need RFC 2822."""
    if not value:
        return "Tue, 01 Jan 2000 00:00:00 +0000"
    text = str(value).strip()
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        try:
            parsed = parsedate_to_datetime(text)
        except (TypeError, ValueError):
            return text
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return format_datetime(parsed)


def convert_to_rfc822(cfg, msg):
    """Converts a Graph message dictionary to RFC822 format."""
    try:
        subject = msg.get("subject") or "(No Subject)"
        sender = ((msg.get("from") or {}).get("emailAddress") or {}).get("address")
        sender = sender or "unknown@example.com"

        def join_addresses(key):
            return ", ".join(
                ((rec.get("emailAddress") or {}).get("address") or "unknown@example.com")
                for rec in (msg.get(key) or [])
            )

        all_recipients = join_addresses("toRecipients")
        cc_list = join_addresses("ccRecipients")
        bcc_list = join_addresses("bccRecipients")
        if cc_list:
            all_recipients += f", {cc_list}"
        if bcc_list:
            all_recipients += f", {bcc_list}"

        date = _rfc2822_date(msg.get("receivedDateTime"))
        body = extract_email_body(msg, cfg.email_format)
        subtype = "html" if cfg.email_format == "html" else "plain"

        attachments = (msg.get("attachments") or []) if cfg.migrate_attachments else []

        if attachments:
            email_msg = MIMEMultipart()
            email_msg.attach(MIMEText(body, subtype, "utf-8"))
            for attachment in attachments:
                content_bytes = attachment.get("contentBytes")
                if not content_bytes:
                    safe_print(f"Skipping attachment without content: "
                               f"{attachment.get('name')!r}")
                    continue
                try:
                    filename = str(attachment.get("name") or "unknown")
                    filename = filename.replace("\r", "").replace("\n", "").replace('"', "")

                    mime_attachment = MIMEBase("application", "octet-stream")
                    mime_attachment.set_payload(base64.b64decode(content_bytes))
                    encoders.encode_base64(mime_attachment)
                    mime_attachment.add_header(
                        "Content-Disposition", f'attachment; filename="{filename}"')
                    email_msg.attach(mime_attachment)
                except Exception as e:
                    safe_print(f"Error processing attachment {attachment.get('name')!r}: {e}")
        else:
            email_msg = MIMEText(body, subtype, "utf-8")

        email_msg["From"] = sender
        email_msg["To"] = all_recipients
        email_msg["Subject"] = subject
        email_msg["Date"] = date

        return email_msg.as_string()

    except Exception as e:
        safe_print(f"Conversion error for message {msg.get('id')!r}: {e}")
        return None


# Authentication -------------------------------------------------------------

def _create_account(cfg, account_tokens_dir):
    return Account(
        (cfg.client_id, cfg.client_secret),
        tenant_id=cfg.tenant_id,
        token_backend=FileSystemTokenBackend(token_path=account_tokens_dir),
        auth_flow_type="credentials",
    )


def _read_access_token(account, token_path):
    token = account.con.token_backend.get_access_token()
    if token and token.get("secret"):
        return token["secret"]

    # O365 < 2.1 stored a flat JSON dict; keep supporting existing token files.
    if os.path.exists(token_path):
        try:
            with open(token_path, "r") as token_file:
                data = json.load(token_file)
        except (OSError, ValueError):
            return None
        return data.get("access_token")
    return None


def authenticate_account(cfg, target_email, force_reauth=True):
    """Authenticate with app credentials and return (access_token, account)."""
    account_tokens_dir = os.path.join(TOKENS_DIR, f"token_{target_email}")
    token_path = os.path.join(account_tokens_dir, "o365_token.txt")

    if force_reauth and os.path.exists(account_tokens_dir):
        shutil.rmtree(account_tokens_dir)
    os.makedirs(account_tokens_dir, exist_ok=True)

    account = _create_account(cfg, account_tokens_dir)
    try:
        authenticated = account.is_authenticated
    except ValueError:
        # Stale token file written by an older O365 release: start over.
        shutil.rmtree(account_tokens_dir, ignore_errors=True)
        os.makedirs(account_tokens_dir, exist_ok=True)
        account = _create_account(cfg, account_tokens_dir)
        authenticated = False

    if not authenticated:
        safe_print(f"Authenticating against Microsoft Graph for {target_email}.")
        if not account.authenticate():
            raise MigrationError(f"authentication failed for {target_email}")

    access_token = _read_access_token(account, token_path)
    if not access_token:
        raise MigrationError(f"no access token available for {target_email}")

    safe_print(f"Successfully authenticated for {target_email}.")
    return access_token, account


# IMAP target ----------------------------------------------------------------

def build_imap_ssl_context(cfg):
    """TLS context for the target server. Certificates are verified by default."""
    context = ssl.create_default_context()
    if not cfg.imap_ssl_verify:
        context.check_hostname = False
        context.verify_mode = ssl.CERT_NONE
    return context


def connect_to_target_imap(cfg, server, email_user, password=None, timeout=30):
    context = build_imap_ssl_context(cfg)
    try:
        mail = imaplib.IMAP4_SSL(server, timeout=timeout, ssl_context=context)
        mail.login(email_user, password)
        mail.noop()
    except ssl.SSLError as e:
        safe_print(f"TLS error connecting to {server}: {e}")
        return None
    except imaplib.IMAP4.error as e:
        safe_print(f"Failed to connect to {server}: {e}")
        return None
    except OSError as e:
        safe_print(f"Network error connecting to {server}: {e}")
        return None

    verification = "TLS certificate verified" if cfg.imap_ssl_verify else \
        "TLS certificate NOT verified"
    safe_print(f"Successfully connected to {server} as {email_user} ({verification}).")
    return mail


IMAP_LIST_RE = re.compile(r'^\((?P<flags>[^)]*)\)\s+(?P<delim>NIL|"(?:[^"\\]|\\.)*")\s+(?P<name>.+)$')


def get_imap_hierarchy(mail):
    """Returns the IMAP hierarchy delimiter and the existing mailbox names."""
    status, data = mail.list()
    if status != "OK":
        raise MigrationError(f"IMAP LIST failed: {data!r}")

    delimiter = "/"
    names = []
    for item in data or []:
        if item is None:
            continue
        line = item.decode("utf-8", errors="replace") if isinstance(item, bytes) else item
        match = IMAP_LIST_RE.match(line.strip())
        if not match:
            continue
        raw_delimiter = match.group("delim")
        if raw_delimiter != "NIL":
            delimiter = raw_delimiter.strip('"')
        name = match.group("name").strip()
        if len(name) >= 2 and name[0] == '"' and name[-1] == '"':
            name = name[1:-1].replace('\\"', '"')
        names.append(name)
    return delimiter, names


def list_imap_folders(mail):
    """Returns the folder names reported by IMAP LIST."""
    return get_imap_hierarchy(mail)[1]


def imap_utf7_encode(value):
    """Encodes a mailbox name as IMAP modified UTF-7 (RFC 3501 section 5.1.3)."""
    encoded = []
    pending = []

    def flush():
        if pending:
            data = "".join(pending).encode("utf-16-be")
            text = base64.b64encode(data).decode("ascii").replace("/", ",").rstrip("=")
            encoded.append(f"&{text}-")
            pending.clear()

    for char in value:
        if 0x20 <= ord(char) <= 0x7E:
            flush()
            encoded.append("&-" if char == "&" else char)
        else:
            pending.append(char)
    flush()
    return "".join(encoded)


def imap_quote(name):
    """Quotes a mailbox name for CREATE/APPEND commands."""
    return '"' + name.replace("\\", "\\\\").replace('"', '\\"') + '"'


def sanitize_folder_component(name):
    name = re.sub(r"[\x00-\x1f\x7f]", "", str(name)).strip()
    if name in ("", ".", ".."):
        return "Unnamed"
    return name[:100]


def build_target_path(cfg, source_path, delimiter):
    """Target mailbox path for a source folder path (unencoded)."""
    parent = [sanitize_folder_component(part)
              for part in cfg.target_parent_folder.split("/") if part.strip()]
    parts = [sanitize_folder_component(part)
             for part in str(source_path or "Inbox").split("/") if part.strip()]
    if not parts:
        parts = ["Inbox"]

    components = parent + parts
    if delimiter != "/":
        # A name containing the delimiter would accidentally nest on the target.
        components = [part.replace(delimiter, "_") for part in components]
    return delimiter.join(components)


def _create_mailbox(mail, wire_name):
    try:
        status, data = mail.create(imap_quote(wire_name))
    except imaplib.IMAP4.error as e:
        return False, str(e)
    if status == "OK":
        return True, ""

    parts = []
    for item in data or []:
        if isinstance(item, bytes):
            parts.append(item.decode("utf-8", "replace"))
        elif item is not None:
            parts.append(str(item))
    return False, " ".join(parts)


def ensure_target_folder(mail, path, delimiter, state):
    """Creates missing mailboxes (parents first) and returns the wire name."""
    wire_name = imap_utf7_encode(path)
    if wire_name.lower() in state["existing"]:
        return wire_name

    current = ""
    for part in path.split(delimiter):
        current = part if not current else f"{current}{delimiter}{part}"
        wire_part = imap_utf7_encode(current)
        if wire_part.lower() in state["existing"]:
            continue

        created, detail = _create_mailbox(mail, wire_part)
        if created:
            safe_print(f"Created target folder '{current}'")
        elif "exist" not in detail.lower():
            safe_print(f"Could not create target folder '{current}': {detail}")
        # Remember the attempt so every message does not retry CREATE.
        state["existing"].add(wire_part.lower())

    return wire_name


def default_target_folder(display_name, state, warned):
    """Maps a standard Outlook folder to the target, falling back to INBOX."""
    candidate = TARGET_FOLDER_MAP.get(display_name, "INBOX")
    wire_name = imap_utf7_encode(candidate)
    if wire_name.lower() not in state["existing"]:
        if candidate.upper() != "INBOX" and display_name not in warned:
            warned.add(display_name)
            safe_print(f"Target folder '{candidate}' does not exist; "
                       f"'{display_name}' falls back to INBOX")
        wire_name = imap_utf7_encode("INBOX")
    return wire_name


def _append_with_retries(mail, wire_folder, payload, attempts=3):
    for attempt in range(1, attempts + 1):
        try:
            status, response = mail.append(
                imap_quote(wire_folder), None,
                imaplib.Time2Internaldate(time.time()), payload,
            )
        except (OSError, imaplib.IMAP4.abort) as e:
            if attempt == attempts:
                safe_print(f"Append failed after {attempts} attempts: {e}")
                return False
            time.sleep(2)
            continue
        except Exception as e:
            safe_print(f"Append failed: {e}")
            return False

        if status == "OK":
            return True
        safe_print(f"Append to '{wire_folder}' failed: {response}")
        return False
    return False


def migrate_emails(cfg, target_mail, messages, board=None, key=None, session=None):
    """Migrates a list of Graph message dictionaries to the target account."""
    delimiter, existing = get_imap_hierarchy(target_mail)
    state = {"existing": {name.lower() for name in existing}}
    warned = set()
    migrated = failed = 0
    current_folder = None

    for msg in messages:
        ok = False
        display_folder = ""
        try:
            if not isinstance(msg, dict):
                safe_print(f"Unexpected message format: {msg!r}")
            else:
                source_path = msg.get("folderPath") or msg.get("folderName") or "Inbox"

                if cfg.target_parent_folder:
                    target_path = build_target_path(cfg, source_path, delimiter)
                    wire_folder = ensure_target_folder(
                        target_mail, target_path, delimiter, state)
                    display_folder = target_path
                else:
                    display_folder = msg.get("folderName") or "Inbox"
                    wire_folder = default_target_folder(display_folder, state, warned)

                if board is not None and key is not None and display_folder != current_folder:
                    current_folder = display_folder
                    board.set_folder(key, display_folder)

                email_message = convert_to_rfc822(cfg, msg)
                if not isinstance(email_message, str):
                    safe_print(f"Email conversion failed for message: {msg.get('id')}")
                else:
                    ok = _append_with_retries(
                        target_mail, wire_folder, email_message.encode("utf-8"))

        except Exception as e:
            safe_print(f"Error migrating email: {e}")

        if ok:
            migrated += 1
            if session is not None:
                session.record(msg.get("id"), display_folder)
        else:
            failed += 1
        if board is not None and key is not None:
            board.advance(key, ok)

    return migrated, failed


def migrate_mailbox(cfg, mailbox, board=None):
    """Migrates a single mailbox and returns (migrated, failed)."""
    source_email, target_server, target_email, target_password = mailbox
    safe_print(f"[{source_email}] starting migration to {target_email} ({target_server})")

    key = None
    session = None
    try:
        access_token, _ = authenticate_account(cfg, source_email)
        if cfg.restartable:
            if cfg.reset_session:
                MigrationSession(session_path(cfg, mailbox), source_email, target_email,
                                 target_server, cfg.target_parent_folder).reset()
            session = MigrationSession.load(cfg, mailbox)
            safe_print(f"[{source_email}] restartable session {session.path} "
                       f"({len(session.ids)} message(s) already recorded)")

        messages = fetch_all_emails(cfg, source_email, access_token,
                                    skip_ids=session.ids if session is not None else None)
        if not messages:
            if session is not None:
                safe_print(f"[{source_email}] nothing new to migrate "
                           f"({len(session.ids)} message(s) in the session).")
            else:
                safe_print(f"[{source_email}] no messages to migrate.")
            return 0, 0

        safe_print(f"[{source_email}] fetched {len(messages)} message(s); "
                   f"connecting to {target_server}")
        target_mail = connect_to_target_imap(cfg, target_server, target_email, target_password)
        if target_mail is None:
            safe_print(f"[{source_email}] failed to connect to target IMAP server {target_server}")
            return 0, 0

        if board is not None:
            key = board.register(source_email, len(messages))

        try:
            migrated, failed = migrate_emails(cfg, target_mail, messages, board, key,
                                              session=session)
        finally:
            try:
                target_mail.logout()
            except Exception:
                pass
            if session is not None:
                session.close()

        if board is not None and key is not None:
            board.finish(key, "done" if failed == 0 else "done with errors")
        else:
            safe_print(f"[{source_email}] finished: {migrated} migrated, {failed} failed "
                       f"-> {target_email}")
        return migrated, failed

    except Exception as e:
        if board is not None and key is not None:
            board.finish(key, "failed")
        if session is not None:
            session.close()
        safe_print(f"[{source_email}] migration failed: {e}")
        return 0, 0


# Dry run --------------------------------------------------------------------

def _report(label, ok, detail=""):
    status = "OK  " if ok else "FAIL"
    line = f"  {label:<26} {status}"
    if detail:
        line += f" {detail}"
    safe_print(line)


def _sample_message(cfg, access_token, source_email, folders):
    """Read and convert one message to prove read access end-to-end."""
    candidates = [f for f in folders if f["totalItemCount"] > 0]
    if not candidates:
        return None, None
    candidates.sort(key=lambda f: f["displayName"] != "Inbox")

    folder = candidates[0]
    page = graph_get(
        f"{_user_url(source_email)}/mailFolders/{folder['id']}/messages",
        access_token, params={"$top": 1},
    )
    messages = page.get("value") or []
    if not messages:
        return folder, None

    message = messages[0]
    message["folderName"] = folder["displayName"]
    message["folderPath"] = folder["path"]
    if cfg.migrate_attachments:
        message["attachments"] = fetch_attachments(access_token, source_email, message["id"])
    return folder, convert_to_rfc822(cfg, message)


def _describe_filters(cfg):
    parts = []
    parts.append(f"folders={','.join(cfg.migrate_folders)}" if cfg.migrate_folders
                 else "folders=all")
    parts.append(f"limit={cfg.migrate_limit or 'all'}")
    if cfg.migrate_after:
        parts.append(f"after={cfg.migrate_after}")
    if cfg.migrate_before:
        parts.append(f"before={cfg.migrate_before}")
    parts.append(f"target parent={cfg.target_parent_folder or '(none)'}")
    parts.append(f"restartable={'on' if cfg.restartable else 'off'}")
    return ", ".join(parts)


def dry_run(cfg, mailboxes):
    """Verify Graph credentials, mailbox access, IMAP logins. Writes nothing."""
    safe_print("=" * 78)
    safe_print("DRY RUN - no messages will be written to any target mailbox")
    safe_print(f"  body format ......... {cfg.email_format}")
    safe_print(f"  attachments ......... {'on' if cfg.migrate_attachments else 'off'}")
    safe_print(f"  graph page size ..... {cfg.graph_page_size}")
    safe_print(f"  IMAP TLS verify ..... {'on' if cfg.imap_ssl_verify else 'OFF (insecure)'}")
    safe_print(f"  selection ........... {_describe_filters(cfg)}")
    safe_print("=" * 78)

    failures = 0
    for index, (source_email, target_server, target_email, target_password) in \
            enumerate(mailboxes, start=1):
        safe_print(f"\n[{index}/{len(mailboxes)}] {source_email}  ->  {target_email}  "
                   f"@ {target_server}")
        mailbox_ok = True

        access_token = None
        folders = []
        try:
            access_token, _ = authenticate_account(cfg, source_email, force_reauth=False)
            _report("Graph authentication", True, f"tenant {cfg.tenant_id}")
        except Exception as e:
            mailbox_ok = False
            _report("Graph authentication", False, str(e))

        if access_token:
            try:
                all_folders = get_mail_folders(cfg, access_token, source_email)
                folders, missing = select_folders(all_folders, cfg.migrate_folders)
                total = sum(f["totalItemCount"] for f in folders)
                detail = f"{len(folders)} of {len(all_folders)} folder(s)"
                if cfg.migrate_limit:
                    detail += f", up to {cfg.migrate_limit} message(s) (of ~{total})"
                else:
                    detail += f", ~{total} message(s)"
                _report("Mail folders", True, detail)
                for name in missing:
                    _report("  filter", False, f"folder not found: {name!r}")
                    mailbox_ok = False
            except Exception as e:
                mailbox_ok = False
                _report("Mail folders", False, str(e))

        if access_token and folders:
            try:
                folder, rfc822 = _sample_message(cfg, access_token, source_email, folders)
                if folder is None:
                    _report("Sample message", True, "skipped, no messages in selection")
                elif rfc822 is None:
                    mailbox_ok = False
                    _report("Sample message", False, "message conversion failed")
                else:
                    _report("Sample message", True,
                            f"{len(rfc822)} bytes, body={cfg.email_format}, "
                            f"folder='{folder['path']}'")
            except Exception as e:
                mailbox_ok = False
                _report("Sample message", False, str(e))

        target_mail = connect_to_target_imap(cfg, target_server, target_email, target_password)
        if target_mail is None:
            mailbox_ok = False
            _report("IMAP login + TLS", False, target_server)
        else:
            try:
                delimiter, target_names = get_imap_hierarchy(target_mail)
                _report("IMAP login + TLS", True, f"{len(target_names)} folder(s), "
                                                  f"delimiter {delimiter!r}")

                if folders:
                    existing = {name.lower() for name in target_names}
                    if cfg.target_parent_folder:
                        desired = [build_target_path(cfg, f["path"], delimiter)
                                   for f in folders]
                        missing = [path for path in desired
                                   if imap_utf7_encode(path).lower() not in existing]
                        _report("Folder mapping", True,
                                f"tree under '{cfg.target_parent_folder}': "
                                f"{len(desired) - len(missing)} exist, "
                                f"{len(missing)} to create")
                        if missing:
                            preview = ", ".join(missing[:4]) + (" ..." if len(missing) > 4 else "")
                            _report("  will create", True, preview)
                    else:
                        fallback = [
                            f["path"] for f in folders
                            if imap_utf7_encode(TARGET_FOLDER_MAP.get(f["displayName"], "INBOX"))
                            .lower() not in existing
                        ]
                        if fallback:
                            preview = ", ".join(fallback[:5]) + (" ..." if len(fallback) > 5 else "")
                            _report("Folder mapping", True,
                                    f"{len(folders) - len(fallback)}/{len(folders)} mapped, "
                                    f"{len(fallback)} fall back to INBOX ({preview})")
                        else:
                            _report("Folder mapping", True,
                                    f"all {len(folders)} folder(s) mapped")
            except Exception as e:
                mailbox_ok = False
                _report("IMAP folder listing", False, str(e))
            finally:
                try:
                    target_mail.logout()
                except Exception:
                    pass

        if not mailbox_ok:
            failures += 1

    safe_print("\n" + "=" * 78)
    ready = len(mailboxes) - failures
    safe_print(f"Dry run finished: {ready}/{len(mailboxes)} mailbox(es) ready, "
               f"{failures} with problems.")
    safe_print("No messages were written.")
    return 0 if failures == 0 else 1


# Entry point ----------------------------------------------------------------

def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        description="Migrate Microsoft 365 mailboxes to an IMAP server.")
    parser.add_argument("--dry-run", action="store_true",
                        help="verify Graph/IMAP credentials and access, write nothing")
    parser.add_argument("--csv", metavar="PATH",
                        help=f"mailbox CSV file (default: {DEFAULT_CSV})")
    parser.add_argument("--env-file", metavar="PATH",
                        help=f"environment file (default: {DEFAULT_ENV})")
    parser.add_argument("--target-parent-folder", metavar="PATH",
                        help="migrate the source folder tree under this target folder "
                             "(created if missing)")
    parser.add_argument("--folders", metavar="LIST",
                        help="comma-separated source folders or paths (subtree match)")
    parser.add_argument("--limit", type=int, metavar="N",
                        help="migrate at most N messages per mailbox (0 = all)")
    parser.add_argument("-i", "--interactive", action="store_true",
                        help="guided single-mailbox wizard (needs requirements-tui.txt)")
    parser.add_argument("--restartable", action="store_true",
                        help="record migrated message IDs per mailbox and resume later")
    parser.add_argument("--reset-session", action="store_true",
                        help="delete existing restart data before migrating (implies "
                             "--restartable)")
    parser.add_argument("--after", metavar="DATE",
                        help="only messages received on/after DATE "
                             "(YYYY-MM-DD or ISO-8601, UTC)")
    parser.add_argument("--before", metavar="DATE",
                        help="only messages received before DATE "
                             "(YYYY-MM-DD or ISO-8601, UTC)")
    return parser.parse_args(argv)


def main(argv=None):
    global _active_board

    args = parse_args(argv)

    if args.interactive:
        from interactive import run_interactive
        return run_interactive(args)

    try:
        cfg = load_config(args.env_file)
        if args.csv:
            cfg.csv_file = args.csv
        if args.target_parent_folder is not None:
            cfg.target_parent_folder = normalize_parent_folder(args.target_parent_folder)
        if args.folders is not None:
            cfg.migrate_folders = parse_folder_list(args.folders)
        if args.limit is not None:
            if args.limit < 0:
                raise MigrationError("--limit must be 0 or a positive integer")
            cfg.migrate_limit = args.limit
        if args.after is not None:
            cfg.migrate_after = normalize_date_bound(args.after)
        if args.before is not None:
            cfg.migrate_before = normalize_date_bound(args.before)
        if cfg.migrate_after and cfg.migrate_before and cfg.migrate_after >= cfg.migrate_before:
            raise MigrationError("--after must be earlier than --before")
        if args.restartable or args.reset_session:
            cfg.restartable = True
        if args.reset_session:
            cfg.reset_session = True
        mailboxes = read_mailboxes_from_csv(cfg.csv_file)
    except MigrationError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 2

    if not mailboxes:
        print(f"Error: no mailboxes found in {cfg.csv_file}.", file=sys.stderr)
        return 2

    try:
        if args.dry_run:
            return dry_run(cfg, mailboxes)

        board = ProgressBoard()
        _active_board = board
        threads = []
        for mailbox in mailboxes:
            thread = threading.Thread(target=migrate_mailbox, args=(cfg, mailbox, board))
            threads.append(thread)
            thread.start()

        for thread in threads:
            thread.join()
        _active_board = None

        print("All mailboxes have been migrated.")
        return 0
    except KeyboardInterrupt:
        _active_board = None
        print("\nInterrupted.")
        return 130


if __name__ == "__main__":
    sys.exit(main())
