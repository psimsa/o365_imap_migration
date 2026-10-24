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

## Next / open questions

- [ ] Run a real end-to-end trial (`--dry-run`, then `--limit 25`) and tag `1.1.0`.
- [ ] `Bcc` recipients are folded into the `To` header.
- [ ] `INTERNALDATE` is migration time rather than `receivedDateTime`.
- [ ] Original `Message-ID`/`Reply-To` headers are lost during re-rendering.
- [ ] Optionally add `--no-progress` for very quiet logs.
