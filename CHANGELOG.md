# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

Planned release: `1.1.0`. Everything below is on `feature/interactive-tui`
(based on `feature/migration-controls`).

### Added

- **Draft interactive wizard** (`-i` / `--interactive`): single-mailbox flow that picks a
  mailbox from `details.csv` or asks for one, verifies Graph credentials from `.env`
  (hidden input, optional atomic write-back), shows account statistics, asks for folder
  scope and a message limit, validates the IMAP login, offers a root-only target parent
  folder (numbered menu under 10 folders, fuzzy typeahead otherwise, new names are
  created), then shows a read-only validation screen and the final plan. Confirmation
  defaults to **No** and no message is written before an explicit `yes`.
- Optional TUI dependencies in `requirements-tui.txt` (`rich`, `prompt_toolkit`); batch
  mode does not import them.
- Unit-test suite (27 tests, hermetic fake Graph/IMAP backends) and a GitHub Actions
  workflow running it on Python 3.10 and 3.12.
- `--dry-run`: authenticates against Graph, lists source folders, reads and converts one
  real message, logs in to every target IMAP server (TLS verified) and reports the folder
  mapping, folders to create and per-mailbox OK/FAIL. Writes nothing; exit code `0`/`1`.
- `TARGET_PARENT_FOLDER` / `--target-parent-folder`: mirrors the O365 folder tree below a
  target folder, creating missing mailboxes parents-first, using the server's hierarchy
  delimiter and IMAP modified UTF-7 for non-ASCII names.
- `MIGRATE_FOLDERS` / `--folders`: comma-separated source folders or paths; a folder also
  selects its subtree. Unknown folders are reported (dry run fails, migration warns).
- `MIGRATE_LIMIT` / `--limit`: cap the number of messages per mailbox (newest first).
- Live migration progress per mailbox: done/total, percentage, ok/failed counts,
  messages/second, ETA and current folder. Redrawn in place on a TTY, throttled lines
  (~2 s) when redirected; a single renderer keeps thread warnings from corrupting it.
- Per-folder/per-page fetch progress.
- `--csv` and `--env-file` options; `CSV_FILE` setting.
- `CHANGELOG.md`, `.env.example`, `details.csv.example`, `.gitignore`, `TASKS.md`.
- Graph retry/backoff for `429`/`5xx` honouring `Retry-After`.
- RFC 2822 `Date` header conversion from Graph ISO 8601 timestamps.

### Changed

- `load_config()` accepts `require_credentials=False` so the wizard can prompt for missing
  `SOURCE_*` values; `migrate_mailbox()` now returns `(migrated, failed)` for callers.
- Configuration moved from `config.json` to a gitignored `.env` (`python-dotenv`).
  `details.csv` (target passwords) is gitignored as well; templates use the `.example` suffix.
- Dependencies updated: `O365==2.1.10`, `requests==2.34.2`, `python-dotenv==1.2.4`.
  Adapted to the 2.1.x token cache and `requested_scopes` API.
- Graph access now sets `$top` and follows `@odata.nextLink` for folders, messages and
  attachments; nested folders are traversed via `childFolders`.
- Message bodies follow `EMAIL_FORMAT`: `html` keeps the Graph HTML body, `plain` strips it.
- Migration reports a final `migrated, failed` count per mailbox.
- README rewritten around setup, dry run, options, behaviour and limitations.

### Fixed

- Only the first page (10 items) of each folder was migrated; now all pages are read.
- Nested O365 folders were never migrated; now they are traversed (and mirrored when a
  target parent folder is configured).
- `EMAIL_FORMAT` in the config was ignored; bodies were always converted to plain text.
- `Date` headers contained ISO 8601 values instead of RFC 2822.
- The append retry loop could count one failure three times and retried on permanent errors.
- Attachment filenames could inject CR/LF/quotes into `Content-Disposition`.
- CSV parsing now skips blank/malformed rows, trims values and de-duplicates entries.

### Security

- The target IMAP connection verifies the server certificate (`ssl.create_default_context`).
  `IMAP_SSL_VERIFY=false` is an explicit, warned opt-out for self-signed internal servers.
- Credentials (`.env`, `details.csv`) and cached Graph tokens (`tokens/`) are gitignored;
  the README asks for `chmod 600`.
- Required Graph permission reduced to `Mail.Read` (application). `Mail.ReadWrite` and
  `User.Read.All` are documented as unnecessary; the script only reads from O365.

## [1.0.0] - 2024-10-04

### Added

- Initial release: read messages through Microsoft Graph with an Entra app (client
  credentials), rebuild them as RFC 822 and `APPEND` them to an IMAP server listed in
  `details.csv`. Optional attachments, `html`/`plain` body setting, threaded per-mailbox
  migration.

[Unreleased]: https://github.com/psimsa/o365_imap_migration/compare/1.0.0...HEAD
[1.0.0]: https://github.com/psimsa/o365_imap_migration/releases/tag/1.0.0
