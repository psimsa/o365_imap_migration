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
from dataclasses import dataclass
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


# Thread-safe printing -------------------------------------------------------

print_lock = threading.Lock()


def safe_print(message, end="\n"):
    """Thread-safe print function that flushes the output immediately."""
    with print_lock:
        print(message, end=end, flush=True)


def update_progress(message, final=False):
    """Updates progress by safely printing a message."""
    if final:
        safe_print(message)
    else:
        safe_print(message, end="\r")


# Configuration --------------------------------------------------------------

@dataclass
class Config:
    client_id: str
    client_secret: str
    tenant_id: str
    migrate_attachments: bool = False
    email_format: str = "html"
    imap_ssl_verify: bool = True
    graph_page_size: int = 100
    csv_file: str = DEFAULT_CSV


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


def load_config(env_file=None):
    """Load settings from a .env file (and the process environment)."""
    path = env_file or DEFAULT_ENV
    if env_file and not os.path.exists(path):
        raise MigrationError(f"env file not found: {path}")
    if os.path.exists(path):
        # An explicitly requested env file wins over the process environment.
        load_dotenv(path, override=bool(env_file))

    missing = [name for name in REQUIRED_ENV if not (os.getenv(name) or "").strip()]
    if missing:
        hint = "" if os.path.exists(path) else f" (copy {DEFAULT_ENV}.example to {DEFAULT_ENV} first)"
        raise MigrationError("missing required setting(s): " + ", ".join(missing) + hint)

    email_format = (os.getenv("EMAIL_FORMAT") or "html").strip().lower()
    if email_format not in ("html", "plain"):
        raise MigrationError(f"EMAIL_FORMAT must be 'html' or 'plain', got {email_format!r}")

    return Config(
        client_id=os.environ["SOURCE_CLIENT_ID"].strip(),
        client_secret=os.environ["SOURCE_CLIENT_SECRET"].strip(),
        tenant_id=os.environ["SOURCE_TENANT_ID"].strip(),
        migrate_attachments=env_bool("MIGRATE_ATTACHMENTS", False),
        email_format=email_format,
        imap_ssl_verify=env_bool("IMAP_SSL_VERIFY", True),
        graph_page_size=env_int("GRAPH_PAGE_SIZE", 100, 1, 1000),
        csv_file=(os.getenv("CSV_FILE") or DEFAULT_CSV).strip(),
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
            print(f"Warning: unexpected header in {filename}: {header}. Expected: {CSV_HEADER}")

        for line_number, row in enumerate(reader, start=2):
            row = [cell.strip() for cell in row]
            if not any(row):
                continue
            if len(row) != len(CSV_HEADER):
                print(f"Warning: skipping {filename}:{line_number}: expected "
                      f"{len(CSV_HEADER)} columns, got {len(row)}.")
                continue
            source_email, target_server, target_email, target_password = row
            if not all(row):
                print(f"Warning: skipping {filename}:{line_number}: empty value in row.")
                continue
            key = (source_email.lower(), target_server.lower(), target_email.lower())
            if key in seen:
                print(f"Warning: skipping {filename}:{line_number}: duplicate entry for "
                      f"{source_email} -> {target_email}.")
                continue
            seen.add(key)
            mailboxes.append((source_email, target_server, target_email, target_password))
    return mailboxes


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


def graph_get_all(url, access_token, params=None, timeout=60):
    """GET every page of a Graph collection by following @odata.nextLink.

    Graph returns at most 10 items per page by default, so ignoring
    @odata.nextLink silently drops everything after the first page.
    """
    items = []
    next_url, next_params = url, params
    while next_url:
        page = graph_get(next_url, access_token, params=next_params, timeout=timeout)
        items.extend(page.get("value", []))
        next_url = page.get("@odata.nextLink")
        # The nextLink already carries the query string of the original request.
        next_params = None
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
            print(f"Warning: could not list child folders of '{path}': {e}")
            continue
        stack.extend((child, f"{path}/{child.get('displayName') or 'Unknown'}")
                     for child in children)
    return folders


def fetch_attachments(access_token, email_address, message_id):
    """Fetch attachments for a given message."""
    url = (f"{_user_url(email_address)}/messages/{quote(message_id, safe='')}/attachments")
    return graph_get_all(url, access_token)


def fetch_all_emails(cfg, source_email, access_token):
    """Fetch all emails from the source email account."""
    folders = get_mail_folders(cfg, access_token, source_email)
    all_emails = []

    for folder in folders:
        try:
            messages = graph_get_all(
                f"{_user_url(source_email)}/mailFolders/{folder['id']}/messages",
                access_token,
                params={"$top": cfg.graph_page_size},
            )
        except GraphError as e:
            print(f"Failed to fetch emails from folder '{folder['path']}': {e}")
            continue

        for message in messages:
            message["folderName"] = folder["displayName"]
            if cfg.migrate_attachments:
                try:
                    message["attachments"] = fetch_attachments(
                        access_token, source_email, message["id"])
                except GraphError as e:
                    print(f"Failed to fetch attachments for message {message.get('id')}: {e}")
                    message["attachments"] = []

        all_emails.extend(messages)
        update_progress(f"Fetched {len(messages)} emails from {folder['path']}.")

    safe_print("")
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
                    print(f"Skipping attachment without content: {attachment.get('name')!r}")
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
                    print(f"Error processing attachment {attachment.get('name')!r}: {e}")
        else:
            email_msg = MIMEText(body, subtype, "utf-8")

        email_msg["From"] = sender
        email_msg["To"] = all_recipients
        email_msg["Subject"] = subject
        email_msg["Date"] = date

        return email_msg.as_string()

    except Exception as e:
        print(f"Conversion error for message {msg.get('id')!r}: {e}")
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
        print(f"Authenticating against Microsoft Graph for {target_email}.")
        if not account.authenticate():
            raise MigrationError(f"authentication failed for {target_email}")

    access_token = _read_access_token(account, token_path)
    if not access_token:
        raise MigrationError(f"no access token available for {target_email}")

    update_progress(f"Successfully authenticated for {target_email}.")
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
        print(f"TLS error connecting to {server}: {e}")
        return None
    except imaplib.IMAP4.error as e:
        print(f"Failed to connect to {server}: {e}")
        return None
    except OSError as e:
        print(f"Network error connecting to {server}: {e}")
        return None

    verification = "TLS certificate verified" if cfg.imap_ssl_verify else \
        "TLS certificate NOT verified"
    print(f"Successfully connected to {server} as {email_user} ({verification}).")
    return mail


IMAP_LIST_RE = re.compile(r'^\((?P<flags>[^)]*)\)\s+(?P<delim>NIL|"(?:[^"\\]|\\.)*")\s+(?P<name>.+)$')


def list_imap_folders(mail):
    """Returns the folder names reported by IMAP LIST."""
    status, data = mail.list()
    if status != "OK":
        raise MigrationError(f"IMAP LIST failed: {data!r}")

    names = []
    for item in data or []:
        if item is None:
            continue
        line = item.decode("utf-8", errors="replace") if isinstance(item, bytes) else item
        match = IMAP_LIST_RE.match(line.strip())
        if not match:
            continue
        name = match.group("name").strip()
        if len(name) >= 2 and name[0] == '"' and name[-1] == '"':
            name = name[1:-1].replace('\\"', '"')
        names.append(name)
    return names


def get_target_folders():
    """Mapping of Outlook folder names to target IMAP folder names."""
    return dict(TARGET_FOLDER_MAP)


def migrate_emails(cfg, target_mail, messages):
    """Migrates emails from a list of message dictionaries to a target email account."""
    target_folders = get_target_folders()
    total_messages = len(messages)
    migrated = failed = 0

    for idx, msg in enumerate(messages):
        try:
            if not isinstance(msg, dict):
                update_progress(f"Unexpected message format: {msg!r}")
                failed += 1
            else:
                source_folder = msg.get("folderName", "Inbox")
                target_folder_name = target_folders.get(source_folder, "INBOX")

                email_message = convert_to_rfc822(cfg, msg)
                if not isinstance(email_message, str):
                    update_progress(f"Email conversion failed for message: {msg.get('id')}")
                    failed += 1
                else:
                    email_message_bytes = email_message.encode("utf-8")
                    for attempt in range(3):
                        try:
                            status, response = target_mail.append(
                                target_folder_name, None,
                                imaplib.Time2Internaldate(time.time()), email_message_bytes,
                            )
                        except (OSError, imaplib.IMAP4.abort) as e:
                            if attempt == 2:
                                update_progress(f"Error appending message: {e}")
                                failed += 1
                            else:
                                time.sleep(2)
                            continue
                        except Exception as e:
                            update_progress(f"Error appending message: {e}")
                            failed += 1
                            break

                        if status != "OK":
                            update_progress(f"Failed to append message to {target_folder_name}: "
                                            f"{response}")
                            failed += 1
                        else:
                            migrated += 1
                        break

            progress_percentage = (idx + 1) / total_messages * 100
            update_progress(f"Migrating Emails: {progress_percentage:.2f}% ({idx + 1}/{total_messages})")

        except Exception as e:
            failed += 1
            update_progress(f"Error migrating email: {e}")

    return migrated, failed


def migrate_mailbox(cfg, mailbox):
    """Migrates a single mailbox."""
    source_email, target_server, target_email, target_password = mailbox
    update_progress(f"Starting migration from {source_email} to {target_email}.")

    try:
        access_token, _ = authenticate_account(cfg, source_email)
        safe_print("")
        messages = fetch_all_emails(cfg, source_email, access_token)
        if not messages:
            update_progress(f"No messages to migrate from {source_email}.", final=True)
            return

        target_mail = connect_to_target_imap(cfg, target_server, target_email, target_password)
        if target_mail is None:
            update_progress(f"Failed to connect to target IMAP server for {target_email}.",
                            final=True)
            return

        try:
            migrated, failed = migrate_emails(cfg, target_mail, messages)
        finally:
            try:
                target_mail.logout()
            except Exception:
                pass

        safe_print("")
        update_progress(f"Migration finished for {source_email} -> {target_email}: "
                        f"{migrated} migrated, {failed} failed.", final=True)

    except Exception as e:
        safe_print("")
        update_progress(f"Migration failed for {source_email}: {e}", final=True)


# Dry run --------------------------------------------------------------------

def _report(label, ok, detail=""):
    status = "OK  " if ok else "FAIL"
    line = f"  {label:<26} {status}"
    if detail:
        line += f" {detail}"
    print(line)


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
    if cfg.migrate_attachments:
        message["attachments"] = fetch_attachments(access_token, source_email, message["id"])
    return folder, convert_to_rfc822(cfg, message)


def dry_run(cfg, mailboxes):
    """Verify Graph credentials, mailbox access, IMAP logins. Writes nothing."""
    print("=" * 78)
    print("DRY RUN - no messages will be written to any target mailbox")
    print(f"  body format ......... {cfg.email_format}")
    print(f"  attachments ......... {'on' if cfg.migrate_attachments else 'off'}")
    print(f"  graph page size ..... {cfg.graph_page_size}")
    print(f"  IMAP TLS verify ..... {'on' if cfg.imap_ssl_verify else 'OFF (insecure)'}")
    print("=" * 78)

    failures = 0
    for index, (source_email, target_server, target_email, target_password) in \
            enumerate(mailboxes, start=1):
        print(f"\n[{index}/{len(mailboxes)}] {source_email}  ->  {target_email}  @ {target_server}")
        mailbox_ok = True

        access_token = None
        folders = []
        try:
            access_token, _ = authenticate_account(cfg, source_email, force_reauth=False)
            safe_print("")
            _report("Graph authentication", True, f"tenant {cfg.tenant_id}")
        except Exception as e:
            mailbox_ok = False
            _report("Graph authentication", False, str(e))

        if access_token:
            try:
                folders = get_mail_folders(cfg, access_token, source_email)
                total = sum(f["totalItemCount"] for f in folders)
                _report("Mail folders", True, f"{len(folders)} folder(s), ~{total} message(s)")
            except Exception as e:
                mailbox_ok = False
                _report("Mail folders", False, str(e))

        if access_token and folders:
            try:
                folder, rfc822 = _sample_message(cfg, access_token, source_email, folders)
                if folder is None:
                    _report("Sample message", True, "skipped, no messages in mailbox")
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
                target_names = list_imap_folders(target_mail)
                _report("IMAP login + TLS", True, f"{len(target_names)} folder(s)")

                if folders:
                    mapping = get_target_folders()
                    fallback = [
                        f["path"] for f in folders
                        if mapping.get(f["displayName"], "INBOX") not in target_names
                    ]
                    if fallback:
                        preview = ", ".join(fallback[:5]) + (" ..." if len(fallback) > 5 else "")
                        _report("Folder mapping", True,
                                f"{len(folders) - len(fallback)}/{len(folders)} mapped, "
                                f"{len(fallback)} fall back to INBOX ({preview})")
                    else:
                        _report("Folder mapping", True, f"all {len(folders)} folder(s) mapped")
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

    print("\n" + "=" * 78)
    ready = len(mailboxes) - failures
    print(f"Dry run finished: {ready}/{len(mailboxes)} mailbox(es) ready, "
          f"{failures} with problems.")
    print("No messages were written.")
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
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)

    try:
        cfg = load_config(args.env_file)
        if args.csv:
            cfg.csv_file = args.csv
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

        threads = []
        for mailbox in mailboxes:
            thread = threading.Thread(target=migrate_mailbox, args=(cfg, mailbox))
            threads.append(thread)
            thread.start()

        for thread in threads:
            thread.join()

        print("All mailboxes have been migrated.")
        return 0
    except KeyboardInterrupt:
        print("\nInterrupted.")
        return 130


if __name__ == "__main__":
    sys.exit(main())
