# o365_imap_migration

Migrate users and shared mailboxes from Microsoft 365 to another mail server over
IMAP. It authenticates with an Entra app (client credentials), reads mail through
Microsoft Graph, and appends each message to the target IMAP account.

## Steps

1. Install the requirements.

   ```bash
   python -m pip install -r requirements.txt
   ```

2. Register an app in the Microsoft Entra admin center and grant **application**
   permissions (not delegated) with admin consent:

   - `Mail.Read`
   - `Mail.ReadWrite`
   - `User.Read.All`

   Then create a client secret under **Certificates & secrets → Client secrets**.

3. Copy the config templates and fill them in. Both real files are gitignored.

   ```bash
   cp .env.example .env
   cp details.csv.example details.csv
   ```

   `.env` holds the app credentials, attachment/body-format options, TLS
   verification and paging settings. `details.csv` holds one row per mailbox:

   ```csv
   source_email,target_server,target_email,target_password
   user@contoso.com,imap.example.com,user@example.com,secret
   ```

4. Verify everything before copying anything:

   ```bash
   python o365_imap_migration.py --dry-run
   ```

   The dry run authenticates against Graph, lists the source folders, reads and
   converts one message, logs in to each target IMAP server, and reports which
   folders will map where. Nothing is written.

5. Run the migration.

   ```bash
   python o365_imap_migration.py
   ```

## Options

| Option | Description |
| --- | --- |
| `--dry-run` | Verify credentials and access only; writes nothing |
| `--csv PATH` | Mailbox CSV file (default: `details.csv`, or `CSV_FILE` in `.env`) |
| `--env-file PATH` | Environment file (default: `.env`) |

## Notes

- **Paging:** Graph returns 10 items per page by default. The script sets
  `$top` (`GRAPH_PAGE_SIZE`) and follows `@odata.nextLink` until every folder and
  message has been read.
- **Body format:** Graph only returns HTML bodies. `EMAIL_FORMAT=html` stores the
  HTML as-is; `EMAIL_FORMAT=plain` strips the markup. Some target servers do not
  render HTML well.
- **Folders:** top-level and nested folders are read. `Inbox`, `Drafts`,
  `Sent Items`, `Deleted Items` and `Junk Email` map to their usual IMAP names;
  any other folder falls back to `INBOX`, and the dry run reports the fallbacks.
- **TLS:** the target server's certificate is verified by default. Setting
  `IMAP_SSL_VERIFY=false` disables verification and should only be used for
  internal servers with self-signed certificates.
- **Tokens:** Graph tokens are cached under `tokens/token_<email>/` (gitignored).
  Migration re-authenticates on every run; the dry run reuses a cached token.
- **Secrets:** `.env` and `details.csv` contain credentials. Keep them out of git
  and restrict file permissions (`chmod 600 .env details.csv`).
