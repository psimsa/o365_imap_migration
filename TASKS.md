# Task tracking

Branch: `feature/migration-controls`
Remote: `origin` -> `git@github.com:psimsa/o365_imap_migration.git` (fork)

## Done

- [x] Verify target IMAP TLS certificates by default (`IMAP_SSL_VERIFY` opt-out).
- [x] Move configuration from `config.json` to a gitignored `.env`.
- [x] Rename `details.csv` to `details.csv.example`; gitignore `details.csv`/`tokens/`.
- [x] Follow `@odata.nextLink` for folders, messages and attachments; recurse `childFolders`.
- [x] Honor `EMAIL_FORMAT` (html/plain) and normalize the `Date` header.
- [x] Add `--dry-run` credential/access verification with a report.
- [x] Update dependencies (O365 2.1.10, requests 2.34.2, python-dotenv 1.2.4), adapt to the
      2.1.x token cache and `requested_scopes` API.
- [x] `TARGET_PARENT_FOLDER` / `--target-parent-folder`: mirror the O365 folder tree below a
      target parent folder, creating missing folders (parents first) with the server's
      hierarchy delimiter and IMAP modified UTF-7 names.
- [x] `MIGRATE_FOLDERS` / `--folders`: comma-separated source folders, subtree match.
- [x] `MIGRATE_LIMIT` / `--limit`: cap messages per mailbox (newest first).
- [x] Live progress while migrating: per-mailbox done/total, %, ok/failed, msg/s, ETA and
      current folder; in-place block on a TTY, throttled lines otherwise, and a single
      renderer so thread warnings cannot corrupt it.
- [x] Review documented Graph permissions; README now asks for `Mail.Read` only.
- [x] Point `origin` at the fork `git@github.com:psimsa/o365_imap_migration.git`.
- [x] Add `CHANGELOG.md` and link it from the README.

## Verification

- Local venv `.venv` created with CPython 3.14.6 (via `uv`); all pinned dependencies
  install and the test suite passes on 3.14.6 and on the system 3.12.3. CI matrix now
  covers 3.10, 3.12 and 3.14.
- Real tenant: Graph app credentials from `.env` verified; `pavel@simsa.cz` is readable
  (18 folders, ~50 906 messages) and a sample message with conversion was fetched
  successfully. Target IMAP not tested yet (no credentials).
- Attachments now default to on (`MIGRATE_ATTACHMENTS=true`); `EMAIL_FORMAT` stays `html`.
- `python -m py_compile o365_imap_migration.py`
- Stub harness (24 checks): Graph paging/limits, folder filters and subtree matching,
  parent-folder creation + UTF-7 wire names, TLS `CERT_REQUIRED`/opt-out, default-mapping
  INBOX fallback, append failure counting, progress board (TTY and non-TTY), `safe_print`
  routing, dry run (default/parent/failure), `.env` parsing/validation, CSV dedupe, CLI.
- A real Graph/IMAP run is still required in an environment with the dependencies installed
  (no `pip` available while this was developed). Start with `--dry-run`.

## Investigation: Graph permissions

- `GET /users/{id | userPrincipalName}/mailFolders`, `/messages`, `/messages/{id}/attachments`
  list application `Mail.Read` (and higher `Mail.ReadWrite`) as sufficient.
- `Mail.Read` application = "Read mail in all mailboxes" - covers every Graph call the script
  makes (it only ever reads from Graph).
- `Mail.ReadWrite` would only be needed to modify the source mailbox (mark read/move/delete),
  which the script does not do.
- `User.Read.All` is not needed: the CSV supplies the UPN and the mail endpoints accept
  `/users/{userPrincipalName}/...` directly.
- Conclusion: README requests `Mail.Read` only.

## Restart + time filters (new)

- [x] `--restartable` / `RESTARTABLE`: per-mailbox JSONL session of successfully
      appended message IDs, skipped before `--limit`; header validates source, target,
      server and parent folder; writes are flushed and fsynced periodically; mode `0600`.
- [x] `--reset-session` / `RESET_SESSION` to delete the session and start over.
- [x] `--after` / `--before` (`MIGRATE_AFTER` / `MIGRATE_BEFORE`) with server-side
      `receivedDateTime` filtering, UTC normalization and range validation.
- [x] Interactive wizard: date-filter prompts and a resume prompt (uses or deletes an
      existing session, defaults on when one exists), plus resume/time-filter rows in the plan.
- [x] Tests for date parsing/filtering, skip-before-limit paging, session round-trip,
      mismatch detection, recording of successes and the interactive flow.
- [x] Real-tenant check of the date filter: `before 2020-01-01`, `after 2025-01-01` and a
      2020-2021 range all returned matching messages from `pavel@simsa.cz`.
- [ ] Real-tenant trial of resume + a full filtered migration (needs IMAP credentials).

## Next / open questions

- [ ] Run a real end-to-end trial (`--dry-run`, then `--limit 25`) and tag `1.1.0`.
      Needs the source O365 mailbox address and IMAP server/password for
      `pavel@simsa.cz` (not provided yet).
- [ ] `Bcc` recipients are folded into the `To` header.
- [ ] `INTERNALDATE` is migration time rather than `receivedDateTime`.
- [ ] Original `Message-ID`/`Reply-To` headers are lost during re-rendering.
- [ ] Optionally add `--no-progress` for very quiet logs.

## Interactive mode (`-i`) - draft for review

Spec: [docs/interactive-tui-spec.md](docs/interactive-tui-spec.md) (implemented as draft)

- [x] Optional `--interactive` / `-i` wizard for a single mailbox.
- [x] Verify (and optionally persist) Graph credentials from `.env`.
- [x] Account statistics + migration plan preview.
- [x] Folder / message-limit scope selection.
- [x] IMAP credential validation and root-only parent-folder picker with fuzzy input.
- [x] Read-only validation screen (Graph creds, mailbox, sample message, IMAP, folder
      plan) instead of a dry-run action; migration only after an explicit `yes`
      (confirmation defaults to No).
- [x] Optional CSV row append; optional TUI dependencies (`requirements-tui.txt`).
- [x] Hermetic unit tests (27) and GitHub Actions workflow (Python 3.10/3.12).
- [ ] User testing and feedback on the draft; then finalise prompts/flow and release.

Draft deviations from the spec: folder selection uses numbered inputs rather than
full-screen checkbox dialogs (SSH/narrow-terminal friendly); the parent-folder typeahead
uses `prompt_toolkit`'s fuzzy completer. `PromptToolkitPrompter`/`RichRenderer` need a
manual TTY test, unit tests use `ScriptedPrompter` + `PlainRenderer`.
