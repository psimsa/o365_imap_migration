# Task tracking

Branch: `feature/migration-controls`

## Done

- [x] Verify target IMAP TLS certificates by default (`IMAP_SSL_VERIFY` opt-out).
- [x] Move configuration from `config.json` to a gitignored `.env`.
- [x] Rename `details.csv` to `details.csv.example`; gitignore `details.csv`/`tokens/`.
- [x] Follow `@odata.nextLink` for folders, messages and attachments; recurse `childFolders`.
- [x] Honor `EMAIL_FORMAT` (html/plain) and normalize the `Date` header.
- [x] Add `--dry-run` credential/access verification with a report.
- [x] Update dependencies (O365 2.1.10, requests 2.34.2, python-dotenv 1.2.4), adapt to the
      2.1.x token cache and `requested_scopes` API.

## In progress

- [ ] `TARGET_PARENT_FOLDER` / `--target-parent-folder`: migrate the O365 folder tree below a
      target parent folder, creating missing folders on the IMAP server (hierarchy delimiter,
      modified UTF-7, recursive `CREATE`).
- [ ] `MIGRATE_FOLDERS` / `--folders`: comma-separated source folders (subtree match).
- [ ] `MIGRATE_LIMIT` / `--limit`: cap the number of messages per mailbox.
- [ ] Live progress while migrating (per mailbox: done/total, %, ok/failed, rate, ETA,
      current folder); TTY in-place board and log-friendly lines otherwise.
- [ ] Review the documented Graph permissions and fix the README if `Mail.Read` is enough.

## Investigation: Graph permissions

- `GET /users/{id | userPrincipalName}/mailFolders`, `/messages`, `/messages/{id}/attachments`
  list application `Mail.Read` (and higher `Mail.ReadWrite`) as sufficient.
- `Mail.Read` application = "Read mail in all mailboxes" - covers every Graph call the script
  makes (it only ever reads from Graph).
- `Mail.ReadWrite` would only be needed to modify the source mailbox (mark read/move/delete),
  which the script does not do.
- `User.Read.All` is not needed: the CSV supplies the UPN and the mail endpoints accept
  `/users/{userPrincipalName}/...` directly.
- Conclusion: README should request `Mail.Read` only.

## Open questions

- Should unmapped folders in default mode be created instead of falling back to `INBOX`?
  (Now solved by `TARGET_PARENT_FOLDER`, which mirrors the tree and creates folders.)
- `Bcc` recipients are currently folded into the `To` header (inherited behaviour).
- `INTERNALDATE` is set to migration time rather than `receivedDateTime`.
