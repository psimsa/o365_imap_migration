"""Shared test helpers: dependency stubs and fake Graph/IMAP backends."""

import base64
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


def attachment(index, name="a.txt", content=b"content", omit_content=False):
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "id": f"att-{index}",
        "name": name,
        "size": len(content),
        "contentType": "application/octet-stream",
        "isInline": False,
        **({"contentBytes": base64.b64encode(content).decode()}
           if not omit_content else {}),
    }


class FakeRequests:
    """Graph fake: 4 folders (one nested), 3 messages per folder, 2 pages each.

    Attachment behaviour: message-list requests with ``$expand`` return each
    message with an ``attachments`` array (fileAttachment with contentBytes,
    unless the attachment id is in ``omit_content_for``). Content for
    ``$value`` batch requests comes from ``attachment_content``.
    """

    class RequestException(Exception):
        pass

    calls = []
    message_pages = None      # optional override: {folder_id: [page, ...]}
    omit_content_for = set()  # attachment ids that come back without contentBytes
    attachment_content = {}   # attachment id -> raw bytes (for $value)
    fail_pages_with_top_above = None  # simulate Graph 504 for big pages
    _last_messages_request_had_expand = False

    @classmethod
    def _with_attachments(cls, msg, params):
        """Attaches the inline attachments when the request asks for $expand."""
        if params.get("$expand"):
            msg = dict(msg)
            msg_id = int(str(msg["id"])[1:])
            msg["attachments"] = [
                attachment(msg_id * 100 + index,
                           name=f"file-{msg['id']}-{index}.txt",
                           content=f"content-{msg['id']}-{index}".encode(),
                           omit_content=(f"att-{msg_id * 100 + index}"
                                         in cls.omit_content_for))
                for index in range(1)
            ]
        return msg

    @classmethod
    def get(cls, url, headers=None, params=None, timeout=None):
        params = dict(params or {})
        cls.calls.append((url, params))

        if cls.fail_pages_with_top_above is not None:
            top = params.get("$top")
            if top is not None and int(top) > cls.fail_pages_with_top_above:
                return FakeResponse({"error": {"code": "GatewayTimeout"}}, status=504)

        # Single attachment resource (itemAttachment expansion)
        if "/attachments/" in url and not url.endswith("/attachments"):
            attachment_id = url.rsplit("/attachments/", 1)[1].split("?")[0]
            if attachment_id.startswith("att-item"):
                return FakeResponse({
                    "@odata.type": "#microsoft.graph.itemAttachment",
                    "id": attachment_id,
                    "name": "nested.eml",
                    "item": message(99),
                })
            return FakeResponse({
                "@odata.type": "#microsoft.graph.fileAttachment",
                "id": attachment_id,
                "name": "file.txt",
                "contentBytes": "ZmlsZQ==",
            })

        if url.endswith("/attachments"):
            return FakeResponse({"value": [
                {"@odata.type": "#microsoft.graph.fileAttachment",
                 "id": "att-1", "name": "a.txt", "contentBytes": "aGk="}]})

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
            if "/next/" not in url:
                cls._last_messages_request_had_expand = bool(params.get("$expand"))
            if "/next/" in url:
                # nextLink responses carry the query in the URL; the production
                # code sends no params for them, so the fake must not rely on
                # params here. Attachments are attached only when the original
                # request asked for $expand (recorded per call sequence).
                msg = message(3)
                if cls._last_messages_request_had_expand:
                    msg["attachments"] = [
                        attachment(300 + index,
                                   name=f"file-m3-{index}.txt",
                                   content=f"content-m3-{index}".encode(),
                                   omit_content=(f"att-{300 + index}"
                                                 in cls.omit_content_for))
                        for index in range(1)
                    ]
                return FakeResponse({"value": [msg]})
            if cls.message_pages is not None:
                # Deterministic per-folder message lists (for verification tests).
                folder_id = url.rsplit("/mailFolders/", 1)[1].split("/")[0]
                pages = cls.message_pages.get(folder_id, [[message(1)]])
                if url.startswith("https://graph.test/next/"):
                    index = int(url.rsplit("/", 1)[1].split("?")[0].replace("p", ""))
                else:
                    index = 0
                index = min(index, len(pages) - 1)
                page = [cls._with_attachments(m, params) for m in pages[index]]
                response = {"value": page}
                if index < len(pages) - 1:
                    response["@odata.nextLink"] = (
                        f"https://graph.test/next/{folder_id}?p={index + 1}")
                return FakeResponse(response)
            page = [message(1), message(2)]
            if params.get("$top") == 1:
                page = page[:1]
            page = [cls._with_attachments(m, params) for m in page]
            from urllib.parse import urlencode
            next_url = "https://graph.test/next/messages"
            if params:
                next_url += "?" + urlencode(params)
            return FakeResponse({"value": page, "@odata.nextLink": next_url})
        raise AssertionError("unexpected Graph URL: " + url)

    @classmethod
    def post(cls, url, json=None, headers=None, timeout=None):
        """$batch endpoint: dispatches inner GETs and wraps the bodies."""
        cls.calls.append((url, {"$batch": True}))
        responses = []
        for inner in (json or {}).get("requests", []):
            inner_url = "https://graph.test" + inner["url"]
            if inner["url"].endswith("/$value"):
                attachment_id = inner["url"].rsplit("/attachments/", 1)[1] \
                    .rsplit("/$value", 1)[0]
                # Unknown attachments (the broken inline images) return empty
                # content, exactly like the real Graph endpoint does.
                content = cls.attachment_content.get(attachment_id, b"")
                import base64 as _b64
                responses.append({
                    "id": inner["id"], "status": 200,
                    "body": _b64.b64encode(content).decode(),
                })
            else:
                resp = cls.get(inner_url)
                responses.append({
                    "id": inner["id"], "status": resp.status_code,
                    "body": resp.json(),
                })
        return FakeResponse({"responses": responses})


class FakeIMAP:
    """IMAP fake with UIDPLUS: APPEND returns APPENDUID and stores CRLF bytes."""

    instances = []
    fail_login = False
    capabilities = ("IMAP4rev1", "UIDPLUS")
    uidvalidity = 1

    def __init__(self, host="", port=None, *, ssl_context=None, timeout=None):
        self.host = host
        self.ssl_context = ssl_context
        self.appended = []      # (folder, payload as sent)
        self.created = []
        self._stored = {}       # (folder, uid) -> payload with CRLF endings
        self._next_uid = 1
        self._selected = None
        FakeIMAP.instances.append(self)

    def login(self, user, password):
        if FakeIMAP.fail_login:
            raise mig.imaplib.IMAP4.error("bad credentials")
        self.user = user

    def noop(self):
        return ("OK", [b""])

    def enable(self, *capabilities):
        return ("OK", [b" ".join(c.encode() for c in capabilities)])

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
        uid = str(self._next_uid)
        self._next_uid += 1
        # The migration already sends CRLF; a real server may additionally
        # normalize (e.g. add a trailing CRLF), which normalize_message_bytes
        # tolerates. Store as-is to model that.
        self._stored[(unquoted, uid)] = msg
        self.appended.append((unquoted, msg))
        return ("OK", [f"APPENDUID {self.uidvalidity} {uid}".encode()])

    def select(self, mailbox="INBOX", readonly=False):
        unquoted = mailbox[1:-1] if mailbox.startswith('"') else mailbox
        self._selected = unquoted
        count = sum(1 for folder, _ in self._stored if folder == unquoted)
        return ("OK", [str(count).encode()])

    def response(self, key):
        if key == "UIDVALIDITY":
            return ("OK", [str(self.uidvalidity).encode()])
        return ("OK", [None])

    def uid(self, command, *args):
        if command != "FETCH":
            raise mig.imaplib.IMAP4.error(f"unsupported UID command {command}")
        uid = str(args[0])
        payload = self._stored.get((self._selected, uid))
        if payload is None:
            return ("OK", [b""])
        header = f"{uid} (UID {uid} BODY[] {{{len(payload)}}}".encode()
        return ("OK", [(header, payload), b")"])

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
    FakeRequests.message_pages = None
    FakeRequests.omit_content_for = set()
    FakeRequests.attachment_content = {}
    FakeRequests.fail_pages_with_top_above = None
    FakeRequests._last_messages_request_had_expand = False
    FakeIMAP.instances = []
    FakeIMAP.fail_login = False
    mig.requests = FakeRequests
    mig.imaplib = FakeImaplib
    # The production code uses per-thread keep-alive sessions; the fake acts as one.
    mig._get_session = lambda: FakeRequests
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
        verify_after_migration=True,
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
