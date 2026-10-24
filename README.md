# o365_imap_migration

> Based on the original repo
> [ArunasVysn/o365_imap_migration](https://github.com/ArunasVysn/o365_imap_migration)
> by [ArunasVysn](https://github.com/ArunasVysn).

Migrate users and shared mailboxes from Microsoft 365 to another mail server over
IMAP. It authenticates with an Entra app (client credentials), reads mail through
Microsoft Graph, and appends each message to the target IMAP account.

## Requirements

- Python 3.10+ (tested with CPython 3.10, 3.12 and 3.14)
- A virtual environment is recommended:
  `python3 -m venv .venv && source .venv/bin/activate`
- `python -m pip install -r requirements.txt`
- Optional, for the interactive wizard: `python -m pip install -r requirements-tui.txt`
- An Entra app registration with **application** permissions (not delegated) and
  admin consent:

  | Permission | Needed? | Why |
  | --- | --- | --- |
  | `Mail.Read` | **yes** | "Read mail in all mailboxes" - covers every Graph call this script makes (`/users/{upn}/mailFolders`, `/messages`, `/messages/{id}/attachments`). |
  | `Mail.ReadWrite` | no | Only needed to modify the source mailbox (mark read, move, delete). The script never writes to O365; it copies out via `GET` and writes to the IMAP target. |
  | `User.Read.All` | no | The CSV already provides the UPN and the mail endpoints accept `/users/{userPrincipalName}/...`. The script never calls `/users` or reads profiles. |

  Add `Mail.ReadWrite` only if you later extend the script to clean up the source
  mailbox after copying. To limit the app to specific mailboxes instead of the
  whole tenant, configure an
  [application access policy](https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access).

- A client secret under **Certificates & secrets → Client secrets**.

References: [List messages permissions](https://learn.microsoft.com/graph/api/user-list-messages?view=graph-rest-1.0#permissions),
[List mailFolders](https://learn.microsoft.com/graph/api/user-list-mailfolders?view=graph-rest-1.0),
[Mail.Read](https://learn.microsoft.com/graph/permissions-reference#mailread).

## Setup

```bash
cp .env.example .env                  # app credentials and options
cp details.csv.example details.csv    # one row per mailbox
chmod 600 .env details.csv
```

Both real files are gitignored. `details.csv` columns:

```csv
source_email,target_server,target_email,target_password
user@contoso.com,imap.example.com,user@example.com,secret
```

## Verify before migrating

```bash
python o365_imap_migration.py --dry-run
```

The dry run authenticates against Graph, lists the selected source folders, reads
and converts one real message, logs in to each target IMAP server (TLS verified)
and reports the folder mapping / which folders would be created. Nothing is
written. Exit code is `0` when every mailbox is ready, `1` otherwise.

## Run

```bash
python o365_imap_migration.py
```

Examples:

```bash
# Mirror the whole O365 folder tree under "Migration 2026" on the target.
python o365_imap_migration.py --target-parent-folder "Migration 2026"

# Only Inbox, Sent Items and the Archive/2024 subtree, at most 500 messages each.
python o365_imap_migration.py --folders "Inbox,Sent Items,Archive/2024" --limit 500

# Trial run of 25 newest Inbox messages into a scratch folder.
python o365_imap_migration.py --target-parent-folder Scratch --folders Inbox --limit 25
```

## Interactive mode (draft)

```bash
python -m pip install -r requirements-tui.txt
python o365_imap_migration.py -i
```

The wizard handles **one mailbox** per run:

1. pick a row from `details.csv` or enter a source mailbox manually;
2. verify the Graph app credentials from `.env`, prompting (hidden input) if they are
   missing or rejected, with an optional, atomic write-back to `.env`;
3. show account statistics (folders, message counts, sizes);
4. ask for the target IMAP credentials and verify the login with TLS — a password
   stored in `details.csv` is used only after an explicit confirmation; for a manually
   entered mailbox it also asks whether to save the row to `details.csv` (written
   before migration, so the restart command works);
5. optionally mirror the folder tree under a root-level target folder: a numbered menu
   when the target has fewer than 10 root folders, otherwise a fuzzy typeahead input;
   a non-matching name is created (surrounding quotes are stripped);
6. ask whether to record progress for resuming an interrupted run (default on when a
   session exists; an existing session can be kept, or deleted and restarted);
7. choose all folders or a subset (a parent also selects its subfolders), an optional
   message limit and an optional received-date range;
8. show a **validation screen** (Graph auth, mailbox access, one sample message, IMAP
   login, folder plan, resume state) and the final plan — including the exact
   **restart command** to resume with if the run is interrupted — then ask for
   confirmation. The confirmation defaults to **No** and only an explicit `y`/`yes`
   starts the migration;
9. migrate (streaming), report `migrated`/`failed`, then automatically verify the
   migrated messages on the target (content hash + subject decode) and print the
   restart and re-verify commands.

The validation screen is the interactive dry run: nothing is fetched in bulk and nothing
is written until the explicit confirmation. `--dry-run` is therefore not needed with `-i`
(a note is printed if both are given). `Ctrl-C` aborts without writing.

Draft notes: selection prompts are numbered inputs (not full-screen dialogs) so the wizard
works over SSH and in narrow terminals; the parent-folder input uses `prompt_toolkit`'s
fuzzy completer. Feedback on the flow is welcome before it is finalised.

## Options

| CLI | `.env` | Default | Meaning |
| --- | --- | --- | --- |
| `-i`, `--interactive` | - | off | Guided single-mailbox wizard (needs `requirements-tui.txt`) |
| `--dry-run` | - | off | Verify credentials/access and report; writes nothing |
| `--csv PATH` | `CSV_FILE` | `details.csv` | Mailbox list |
| `--env-file PATH` | - | `.env` | Environment file |
| `--target-parent-folder PATH` | `TARGET_PARENT_FOLDER` | empty | Migrate the source folder tree below this target folder, creating missing folders (parents first). Empty = classic mapping |
| `--folders LIST` | `MIGRATE_FOLDERS` | all | Comma-separated source folders or paths; a folder also selects its subtree |
| `--limit N` | `MIGRATE_LIMIT` | `0` (all) | At most N messages per mailbox, newest first |
| `--restartable` | `RESTARTABLE` | off | Record migrated message IDs per mailbox in `sessions/` and skip them when the run is repeated |
| `--reset-session` | `RESET_SESSION` | off | Delete existing restart data before migrating (implies `--restartable`) |
| `--after DATE` | `MIGRATE_AFTER` | no bound | Only messages received on/after DATE (`YYYY-MM-DD` or ISO-8601, UTC) |
| `--before DATE` | `MIGRATE_BEFORE` | no bound | Only messages received before DATE (exclusive) |
| `--verify` | - | off | Verify messages recorded in restart sessions against the target (no migration) |

Other `.env` settings: `MIGRATE_ATTACHMENTS` (default **on**), `EMAIL_FORMAT`
(`html` by default, or `plain`), `IMAP_SSL_VERIFY`, `GRAPH_PAGE_SIZE`,
`SESSION_DIR`, plus the three `SOURCE_*` credentials.

## Behaviour

- **Progress:** one line per mailbox with `done/total`, percent, `ok/failed`
  counts, messages/second, ETA and the current folder. On a terminal the block is
  redrawn in place; when output is redirected (CI/logs) a line is emitted every
  ~2 seconds and on completion. Fetching also reports per-folder/page progress.
  Warnings are printed without breaking the progress display.
- **Folder tree:** with `TARGET_PARENT_FOLDER` every source path is recreated
  below it (`Migration/Sent Items`, `Migration/Archive/2024`, ...), sanitizing
  invalid characters, using the target server's hierarchy delimiter and IMAP
  modified UTF-7 for non-ASCII names. Without it, `Inbox`, `Drafts`, `Sent Items`,
  `Deleted Items` and `Junk Email` map to their usual IMAP names; unknown folders
  fall back to `INBOX` and the dry run reports them.
- **Attachments:** file attachments are copied as-is. Attached emails
  (`itemAttachment`) are fetched with `$expand=microsoft.graph.itemAttachment/item`
  and re-attached as `message/rfc822` (`.eml`), including their own attachments when
  Graph returns them. Cloud-link attachments become small `.url` files so the link is
  not lost. Messages with attachments Graph exposes no content for are reported, not
  silently dropped.
- **Streaming:** each Graph page is fetched, converted and appended immediately, so
  the target starts receiving mail while the source is still being read and memory
  stays flat (important for 50k-message mailboxes with attachments). The restart
  session file is created when the run starts, not after the first message.
- **Paging:** Graph returns 10 items per page by default. The script sets `$top`
  (`GRAPH_PAGE_SIZE`) and follows `@odata.nextLink` until every folder and
  message has been read; 429/5xx responses are retried honouring `Retry-After`.
- **Time filter:** `--after`/`--before` are applied server-side with
  `receivedDateTime ge/lt ...`, so the limit and paging work on the filtered set.
  Date-only values are UTC midnight; `--before` is exclusive. This is useful for
  backups such as "everything older than 2023-01-01". Folder message counts in the
  plan are collection totals: the date filter is applied when messages are fetched,
  so the final count can be smaller.
- **Parent folder names:** surrounding quotes are stripped, so typing
  `'Office Migration'` (or `"Office Migration"`) targets `Office Migration`.
- **Unicode / encoding:** Graph returns subjects, bodies, attachment names and
  folder names as decoded Unicode; the script re-encodes everything safely for
  the wire:
  - bodies are written as UTF-8 (`charset="utf-8"`, base64) — the whole message
    is pure ASCII, so IMAP `APPEND` never needs `8BITMIME`;
  - non-ASCII subjects become RFC 2047 encoded words and round-trip exactly;
  - non-ASCII attachment filenames become RFC 2231 `filename*=utf-8''...`;
  - embedded `<meta charset=...>` / XML declarations inside HTML bodies are
    rewritten to UTF-8, so mail clients do not mis-decode legacy (e.g.
    windows-1252) declarations;
  - non-ASCII target folder names use IMAP modified UTF-7;
  - messages are sent with CRLF line endings as IMAP requires.
- **Verification:** with `--restartable`, every appended message is recorded with
  its target UID (APPENDUID/UIDPLUS), a SHA-256 of the exact bytes sent and the
  exact subject. `python o365_imap_migration.py --verify` re-reads every recorded
  message from the target (read-only `BODY.PEEK[]`), compares the content hash
  (line endings normalized) and checks the subject decodes back to the original
  — this catches any corruption, including mangled Unicode. The interactive
  wizard verifies automatically after migrating (`VERIFY_AFTER_MIGRATION=false`
  to disable). Entries recorded before UID tracking (or after a UIDVALIDITY
  change) are reported as "not verifiable", not as failures. Exit code is `1`
  if any recorded message is missing or mismatched.
- **Restartable:** with `--restartable`, every successfully appended message ID is
  written to `sessions/<source>__<target>__<server>__<hash>.jsonl` (one JSON object
  per line, header validates source/target/parent folder, mode `0600`). A rerun
  loads the file and skips those IDs *before* applying `--limit`, so you keep
  making progress. Failed messages are not recorded. Use `--reset-session` to
  deliberately start over (for example after changing the target parent folder);
  a session that belongs to a different target is refused rather than mixed in.
  Dry runs and the interactive validation step never write session data.
- **Body format:** Graph only returns HTML bodies. `EMAIL_FORMAT=html` stores the
  HTML as-is; `EMAIL_FORMAT=plain` strips the markup. Some target servers do not
  render HTML well.
- **TLS:** the target server's certificate is verified by default. Setting
  `IMAP_SSL_VERIFY=false` disables verification (MITM risk) and is only meant for
  internal servers with self-signed certificates.
- **Tokens:** Graph tokens are cached under `tokens/token_<email>/` (gitignored).
  A normal migration re-authenticates on every run; the dry run reuses a cached,
  unexpired token.

## Known limitations

- `Bcc` recipients are folded into the `To` header.
- `INTERNALDATE` is the migration time, not the original `receivedDateTime`.
- Message bodies are re-rendered from the Graph JSON body; original headers other
  than `From`/`To`/`Subject`/`Date` (e.g. `Message-ID`, `Reply-To`) are not kept.
- A filter on a parent folder selects it and all of its subfolders.

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for the full history. Unreleased highlights:

- Draft interactive wizard (`-i`) with optional `.env` write-back, account stats,
  folder/limit selection, parent-folder picker, validation screen and
  confirmation-default-No.
- Unit tests (`python -m unittest discover -s tests -t .`) and a GitHub Actions
  workflow running them on Python 3.10 and 3.12.

- `Mail.Read` is the only Graph permission required (README previously asked for
  `Mail.ReadWrite` and `User.Read.All` as well).
- New `--target-parent-folder`, `--folders` and `--limit` switches (with `.env`
  equivalents `TARGET_PARENT_FOLDER`, `MIGRATE_FOLDERS`, `MIGRATE_LIMIT`).
- Live per-mailbox progress (percentage, ok/failed, rate, ETA, current folder).
- Paging fix: all messages and nested folders are now migrated, not just the first
  page of top-level folders.
- TLS certificate verification for the target IMAP server.
- `--dry-run` for a full credential/access check without writing anything.

## Project

This project is **based on**
[ArunasVysn/o365_imap_migration](https://github.com/ArunasVysn/o365_imap_migration)
by [ArunasVysn](https://github.com/ArunasVysn), and is maintained independently at
[psimsa/o365_imap_migration](https://github.com/psimsa/o365_imap_migration).
Task tracking for the current work lives in [TASKS.md](TASKS.md).
