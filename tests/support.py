"""Shared test helpers: dependency stubs and fake Graph/IMAP backends."""

import os
import sys
import types


def _install_dependency_stubs():
    """Keep unit tests hermetic: no O365/dotenv install required."""
    try:
        import O365  # noqa: F401
    except Exception:
        package = types.ModuleType("O365")

        class Account:  # pragma: no cover - replaced by patches in tests
            def __init__(self, *args, **kwargs):
                self.args = args
                self.kwargs = kwargs

        class FileSystemTokenBackend:  # pragma: no cover
            def __init__(self, token_path=None, token_filename=None):
                self.token_path = token_path

        package.Account = Account
        package.FileSystemTokenBackend = FileSystemTokenBackend
        sys.modules["O365"] = package

    try:
        import dotenv  # noqa: F401
    except Exception:
        module = types.ModuleType("dotenv")

        def load_dotenv(path=None, override=False):
            if not path or not os.path.exists(path):
                return False
            with open(path, encoding="utf-8") as handle:
                for line in handle:
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    key, value = line.split("=", 1)
                    key = key.strip()
                    value = value.strip().strip('"').strip("'")
                    if override or key not in os.environ:
                        os.environ[key] = value
            return True

        module.load_dotenv = load_dotenv
        sys.modules["dotenv"] = module


_install_dependency_stubs()

import o365_imap_migration as mig  # noqa: E402


class FakeResponse:
    def __init__(self, payload, status=200, headers=None):
        self._payload = payload
        self.status_code = status
        self.headers = headers or {}
        self.text = "" if status == 200 else "boom"

    def json(self):
        return self._payload


def message(index, folder_name="Inbox"):
    return {
        "id": f"m{index}",
        "subject": f"S{index}",
        "from": {"emailAddress": {"address": "a@b.com"}},
        "toRecipients": [{"emailAddress": {"address": "c@d.com"}}],
        "ccRecipients": [],
        "bccRecipients": [],
        "receivedDateTime": "2024-01-02T03:04:05Z",
        "body": {"contentType": "html", "content": "<b>Hello</b> &amp; bye"},
    }


class FakeRequests:
    """Graph fake: 4 folders (one nested), 3 messages per folder, 2 pages each."""

    class RequestException(Exception):
        pass

    calls = []

    @classmethod
    def get(cls, url, headers=None, params=None, timeout=None):
        params = dict(params or {})
        cls.calls.append((url, params))
        if "/attachments" in url:
            return FakeResponse({"value": [{"name": "a.txt", "contentBytes": "aGk="}]})
        if "/childFolders" in url:
            if "/f1/" in url:
                return FakeResponse({"value": [{"id": "f1c", "displayName": "Sub",
                                                "totalItemCount": 1}]})
            return FakeResponse({"value": []})
        if url.endswith("/mailFolders"):
            return FakeResponse({
                "value": [
                    {"id": "f1", "displayName": "Inbox", "totalItemCount": 5},
                    {"id": "f2", "displayName": "Archive", "totalItemCount": 1},
                ],
                "@odata.nextLink": "https://graph.test/next/folders",
            })
        if url == "https://graph.test/next/folders":
            return FakeResponse({"value": [{"id": "f3", "displayName": "Sent Items",
                                            "totalItemCount": 1}]})
        if "/messages" in url:
            if params.get("$top") == 1:
                return FakeResponse({"value": [message(1)]})
            if url == "https://graph.test/next/messages":
                return FakeResponse({"value": [message(3)]})
            return FakeResponse({"value": [message(1), message(2)],
                                 "@odata.nextLink": "https://graph.test/next/messages"})
        raise AssertionError("unexpected Graph URL: " + url)


class FakeIMAP:
    instances = []
    fail_login = False

    def __init__(self, host="", port=None, *, ssl_context=None, timeout=None):
        self.host = host
        self.ssl_context = ssl_context
        self.appended = []
        self.created = []
        FakeIMAP.instances.append(self)

    def login(self, user, password):
        if FakeIMAP.fail_login:
            raise mig.imaplib.IMAP4.error("bad credentials")
        self.user = user

    def noop(self):
        return ("OK", [b""])

    def list(self):
        return ("OK", [b'(\\HasNoChildren) "/" "INBOX"',
                       b'(\\HasNoChildren) "/" "Drafts"',
                       b'(\\HasNoChildren) "/" "Sent"',
                       b'(\\HasNoChildren) "/" "Deleted Items"',
                       b'(\\HasNoChildren) "/" "Spam"'])

    def create(self, name):
        unquoted = name[1:-1] if name.startswith('"') else name
        self.created.append(unquoted)
        return ("OK", [b"created"])

    def append(self, mailbox, flags, when, msg):
        unquoted = mailbox[1:-1] if mailbox.startswith('"') else mailbox
        self.appended.append((unquoted, msg))
        return ("OK", [b""])

    def logout(self):
        return ("BYE", [b""])


class FakeImaplib:
    IMAP4_SSL = FakeIMAP

    class IMAP4:
        error = Exception
        abort = Exception

    Time2Internaldate = staticmethod(lambda t: '"01-Jan-2024 00:00:00 +0000"')


def install_fakes():
    FakeRequests.calls = []
    FakeIMAP.instances = []
    FakeIMAP.fail_login = False
    mig.requests = FakeRequests
    mig.imaplib = FakeImaplib
    return FakeRequests, FakeIMAP


def make_cfg(**overrides):
    base = dict(
        client_id="client-id",
        client_secret="client-secret",
        tenant_id="tenant-id",
        migrate_attachments=False,
        email_format="html",
        imap_ssl_verify=True,
        graph_page_size=100,
        csv_file="details.csv",
        target_parent_folder="",
        migrate_folders=[],
        migrate_limit=0,
    )
    base.update(overrides)
    return mig.Config(**base)


def sample_folders():
    return [
        {"id": "f1", "displayName": "Inbox", "path": "Inbox", "totalItemCount": 5},
        {"id": "f1c", "displayName": "Sub", "path": "Inbox/Sub", "totalItemCount": 1},
        {"id": "f2", "displayName": "Archive", "path": "Archive", "totalItemCount": 1},
        {"id": "f3", "displayName": "Sent Items", "path": "Sent Items", "totalItemCount": 1},
    ]
