# Interactive mode (`--interactive` / `-i`) — specification & implementation plan

Status: **draft implemented** (`interactive.py`) - awaiting user testing/feedback
Branch: `feature/interactive-tui` (based on `feature/migration-controls`)

## 1. Goals and non-goals

**Goals**

- Guided, single-mailbox wizard on top of the existing batch engine.
- Validate/repair the Graph app credentials from `.env`, with optional write-back.
- Show real O365 statistics and a concrete migration plan before writing anything.
- Validate the target IMAP login (TLS verified) before writing anything.
- Let the user pick (or create) a target parent folder for the folder tree.
- Validate everything the migration needs (credentials, mailbox access, sample message,
  target login, folder plan) on the final screen, then require an explicit
  **yes** (default **no**) before any `APPEND`.
- Reuse `fetch_all_emails`, `migrate_emails`, `ProgressBoard`, TLS context and folder
  helpers — the wizard only orchestrates, it does not duplicate migration logic.

**Non-goals**

- Replacing batch mode: without `-i` nothing changes.
- A separate dry-run action in interactive mode. The final validation screen *is* the dry
  run: it proves credentials, access and the plan without fetching or writing everything,
  and the migration only starts after the user explicitly confirms.
- Multiple mailboxes in one wizard run (one mailbox per invocation, by design).
- Persisting IMAP passwords, or remembering answers between runs.
- Editing arbitrary `.env` keys, managing app registrations, or OAuth device flow.
- A full-screen alternate-buffer application: the wizard is a linear scroll-back flow,
  which works over SSH/mosh, in `tmux`, and on Windows terminals.

## 2. Dependencies

Interactive mode needs richer prompts than `input()`:

| Package | Use | Why |
| --- | --- | --- |
| `rich` | panels, tables, spinners, colour | stats table and plan rendering |
| `prompt_toolkit` | password input, radio/checkbox lists, fuzzy autocomplete | non-echoing secrets, multi-select, "whisper" typeahead |

Recommendation: **optional extra**, not a hard dependency:

- new `requirements-tui.txt` (`rich==…`, `prompt_toolkit==…`, pinned).
- imports happen lazily inside `interactive.py`; batch mode never imports them.
- if `-i` is used without them: exit `2` with
  `Interactive mode needs: python -m pip install -r requirements-tui.txt`.

Rationale: headless servers doing batch migrations should not pull in a terminal UI stack.
If you prefer a single `requirements.txt`, it is a one-line change — see "Decisions".

## 3. Entry point

```
python o365_imap_migration.py -i
python o365_imap_migration.py -i --env-file .env.staging --csv details.prod.csv
```

- New argparse option: `--interactive`, `-i`, `store_true`.
- Dispatched in `main()` **after** `--env-file`/`--csv` resolution, **before** batch
  validation of `details.csv` (the wizard may create a row / enter values manually).
- Requires a TTY (`sys.stdin.isatty() and sys.stdout.isatty()`); otherwise exit `2`
  with a hint to use the non-interactive flags.
- Existing switches become **prefilled defaults**, still editable: `--folders`, `--limit`,
  `--target-parent-folder`. `--csv`/`--env-file` choose the files.
- `--dry-run` has no meaning in interactive mode (the wizard always validates first and
  never writes before confirmation). If it is passed together with `-i`, print a one-line
  note and ignore it.
- `Ctrl-C`/`Ctrl-D` at any prompt: abort with `Interrupted, nothing written.` (exit `130`).

## 4. Wizard flow

Eight steps; each prints a step header. No secrets are ever echoed to the screen or logs.

### Step 1 — Source mailbox

- If `details.csv` exists and has parseable rows, show them:
  `❯ user@contoso.com → user@example.com @ imap.example.com`
  plus `Enter a different mailbox…`.
- If no/blank CSV: go straight to manual entry.
- Manual entry: **source O365 address** (required). The target fields are asked in Step 5.
- When an existing row was picked, its target server/address prefill Step 5 and its
  password is never displayed.

### Step 2 — Graph credentials (`SOURCE_*`)

1. Read `.env` (existing `load_config`). If all three keys exist, **verify** them:
   acquire a client-credentials token against the tenant with a throwaway token cache
   (`tokens/_preflight/`, removed afterwards). No mailbox is touched yet.
2. On failure (missing keys, `invalid_client`, `invalid_tenant`, `unauthorized_client`,
   AADSTS error, network): show the Graph error text (truncated, secrets redacted) and
   prompt for **Client ID**, **Client secret** (hidden input), **Tenant ID** (defaults to
   the current value, then `common`).
3. Re-verify. Up to 3 attempts, then offer to abort.
4. After successful verification, if the values differ from `.env`/were entered manually,
   ask: `Save these credentials to .env? [y/N]` (default **No**). Saving is an in-place,
   atomic edit that only touches the three keys.

### Step 3 — Account statistics

- `authenticate_account(cfg, source_email, force_reauth=False)`, then `get_mail_folders`
  (this also proves the mailbox is readable with the app permission; a 404 becomes
  `mailbox not found or not accessible for this app`).
- Print a table: folder path, messages, size (`sizeInBytes` when returned), plus totals.
  Show at most the top 15 folders by message count and `… N more folder(s)`.

```
Account statistics
┌────────────────────────────┬──────────┬─────────┐
│ Folder                     │ Messages │   Size  │
├────────────────────────────┼──────────┼─────────┤
│ Inbox                      │    8 214 │  2.1 GB │
│ Sent Items                 │    3 002 │  0.9 GB │
│ Archive/2024               │      912 │    1 GB │
│ … 4 more folder(s)         │    1 004 │  0.3 GB │
├────────────────────────────┼──────────┼─────────┤
│ Total                      │   13 132 │  4.3 GB │
└────────────────────────────┴──────────┴─────────┘
```

- If the mailbox has no folders, report it and offer to abort.

### Step 4 — Scope

Two questions, defaults from CLI/env:

1. **Folders**: radio `All folders (N)` / `Choose folders…`.
   - "Choose folders" shows a checkbox list of all folder paths (searchable by typing).
     Selecting a parent selects its subtree (same semantics as `--folders`).
2. **Message limit**: `All messages` / `Newest N messages…` (positive integer; empty
   means all).

The effective values are stored in a copy of `Config` (`migrate_folders`, `migrate_limit`).

### Step 5 — Target IMAP

- Fields: server, address, password (hidden). Prefilled from the chosen CSV row when
  available; password always asked (never stored in the config).
- Connect with the existing `connect_to_target_imap` (TLS verified). On failure show the
  error and allow retry/abort. Warn loudly if `IMAP_SSL_VERIFY=false` in `.env`.
- After login, read `get_imap_hierarchy` (delimiter + mailboxes) — used in Step 6.
- If the mailbox was entered manually, ask (after the plan, Step 7) whether to append it to
  `details.csv`; default **No**, and never write the password to logs.

### Step 6 — Target parent folder

Question: `Mirror the folder tree under an existing/new target folder?`

- **No** → default mapping (`Inbox→INBOX`, `Drafts→Drafts`, `Sent Items→Sent`,
  `Deleted Items→Deleted Items`, `Junk Email→Spam`, everything else → `INBOX`).
- **Yes** → choose a **root-level** target folder:
  - Candidate list = mailboxes reported by `LIST` that contain **no hierarchy delimiter**
    (i.e. top-level only, not `Sent/2024`). INBOX and the usual special-use folders are
    shown but marked `(special)`; they are allowed but not preselectable.
  - **If candidates < 10**: numbered menu of candidates plus `Type a new folder name…`
    (and `Back`).
  - **If candidates ≥ 10**: free-text input with fuzzy autocomplete ("whisper" =
    typeahead suggestions while typing, `prompt_toolkit` fuzzy completer). Typing a
    non-matching name is accepted.
  - Values are trimmed; a value containing the hierarchy delimiter is rejected with an
    explanation (single root folder only).
  - If the input matches a candidate (case-insensitive) → existing folder is reused; the
    plan says `reuse existing`. Otherwise → `will be created`.
- Either way the choice is stored in `cfg.target_parent_folder`.

### Step 7 — Validation, final plan and confirmation

This step replaces a separate dry-run action. It performs the read-only checks the
migration depends on (without fetching or writing the mailbox) and shows the resulting
plan. **If any check fails, migration cannot start** — the user can go back or abort.

Read-only checks, using the credentials and target connection already established:

- Graph app token valid for the tenant.
- Mailbox readable and the selected folders resolvable.
- One sample message from the selection can be read and converted (`html`/`plain`,
  attachments setting), proving message/attachment access.
- Target IMAP login and TLS verification (already done in Step 5, shown again).
- Folder plan: how many target folders will be created vs reused, and warnings
  (e.g. `Junk Email → INBOX` when the target has no `Spam` folder).

```
Validation & plan
  ✓ Graph app credentials        tenant contoso.onmicrosoft.com
  ✓ Mailbox access               user@contoso.com
  ✓ Sample message               Inbox, 14.2 KB, body=html
  ✓ Target IMAP login            imap.example.com (TLS verified)
  ✓ Folder plan                  6 to create, 2 reused
  ⚠ Warning                      Junk Email -> INBOX (no Spam folder on target)

  Source          user@contoso.com      13 132 messages in 18 folders
  Target          user@example.com      imap.example.com
  Folders         Inbox, Sent Items, Archive (+ subfolders)
  Limit           500 newest messages per folder selection
  Attachments     on            Body format   html
  Target parent   Migration 2026  (will be created)

Start the migration? [y/N]
```

- The plan never contains the client secret or IMAP password.
- **Default is No.** Only an explicit `y`/`yes` starts the migration; `N`, empty input,
  `Ctrl-C` or `Ctrl-D` → `Aborted, nothing written.` exit `2`.
- If validation found a failure, replace the confirmation prompt with
  `Back to the previous step? [Y/n]` / `Abort [a]`.

### Step 8 — Execute

- Register a `ProgressBoard` (single row) and call `migrate_mailbox(cfg, mailbox, board)`;
  print the final `migrated / failed` summary.
- Exit `0` only when `failed == 0`; exit `1` when any message failed.
- On success, if the mailbox was entered manually and the user agreed in Step 5, append the
  row to `details.csv` (header created if the file is new; file mode `0600`).
- Final line: counts, elapsed time, and the exact command to rerun the same migration
  non-interactively (useful for automation).

## 5. Components

Keep batch mode untouched; add one module, lazily imported.

```
o365_imap_migration.py     # + --interactive/-i in parse_args, dispatch in main()
interactive.py             # wizard only, imports the core module
requirements-tui.txt       # rich, prompt_toolkit (optional)
tests/test_interactive.py  # scripted-prompter tests (pytest)
```

`interactive.py` contents:

| Component | Responsibility |
| --- | --- |
| `Prompter` (protocol) | `text`, `secret`, `select`, `multi_select`, `confirm` |
| `PromptToolkitPrompter` | real implementation (hidden input, radio/checkbox, fuzzy completer) |
| `ScriptedPrompter` (tests) | returns canned answers, records prompts |
| `preflight_credentials(cfg, prompter)` | verify/repair/persist `SOURCE_*` |
| `choose_mailbox(cfg, prompter)` | CSV row picker or manual entry |
| `choose_scope(cfg, folders, prompter)` | folder subset + limit |
| `choose_parent_folder(cfg, delimiter, names, prompter)` | root-only list or fuzzy input |
| `validate(cfg, mailbox, hierarchy)` | read-only checks + one sample message; returns `Validation` |
| `build_plan(cfg, folders, validation, hierarchy)` | **pure** → `Plan` (counts, target paths, warnings) |
| `format_plan(plan, validation)` | rich table / plain text |
| `write_env_values(path, values)` | in-place atomic `.env` update, comments preserved |
| `append_csv_row(path, row)` | optional CSV append, `0600`, header if new |
| `run_interactive(args)` | orchestrates the steps, returns exit code |

Pure functions (`build_plan`, root-folder filtering, fuzzy matching, env-file editing,
confirmation parsing) carry all logic, so the wizard is testable without a terminal.

## 6. Security requirements

- Secrets are read with hidden input (`is_password=True`), never via argv, never echoed,
  never included in plan output, error messages, or the token cache path.
- `.env` write-back: read lines, replace only matching `KEY=` lines, append missing keys,
  keep comments and order, write `path.tmp` then `os.replace`, `chmod 0600` (POSIX).
- Never persist the IMAP password. Never print it on retry.
- The wizard reuses the verified TLS context; `IMAP_SSL_VERIFY=false` produces a visible
  warning in the plan.
- Ctrl-C leaves no partial writes: `.env`/CSV are written atomically, and no IMAP writes
  happen before the confirmation.
- The validation step only performs `GET`/`LIST`/sample read — no `APPEND`, no `CREATE`,
  no source-side changes.

## 7. Error handling

| Situation | Behaviour |
| --- | --- |
| Missing TUI dependency | exit 2, print `pip install -r requirements-tui.txt` |
| No TTY | exit 2, point at non-interactive flags |
| `.env` missing/blank credentials | prompt, retry ≤3, then abort (exit 2) |
| Graph auth rejected | show truncated AADSTS/Graph error, offer re-enter |
| Mailbox 404/forbidden | show error, offer another source address or abort |
| IMAP login/TLS failure | show error, retry or abort (no writes yet) |
| Source folder filter matches nothing | show warning, go back to Step 4 |
| No messages in scope | show 0, offer back/abort (never "succeed" silently) |
| Any validation check fails | block migration, offer back/abort |
| Confirmation not `yes` | `Aborted, nothing written.` exit 2 |
| Migration with message failures | summary + exit 1; details already printed per message |

## 8. Testing plan

Unit tests (no network, no terminal):

1. `build_plan`: totals with/without limit, subtree folder selection, default mapping vs
   parent folder, "create/reuse" counts, warnings for INBOX fallback.
2. Root-folder filtering: excludes names containing the delimiter; marks special-use;
   threshold `>= 10` switches to fuzzy input.
3. Fuzzy matching: case-insensitive match, no-match accepted as new folder, delimiter in
   input rejected.
4. `write_env_values`: replaces existing keys, appends new ones, preserves comments/order
   and file mode, leaves other lines untouched; idempotent.
5. `append_csv_row`: header when new, dedupe of identical rows, `0600`.
6. Confirmation parsing: only `y`/`yes` (case-insensitive) start the migration; empty input
   and everything else abort; default is No.
7. `validate`: returns one entry per check with pass/fail; a failing sample read marks the
   validation failed and blocks execution (assert no `append` calls).
8. Scripted flow with fakes (reuse the stub Graph/IMAP harness): prompt order, validation
   performed before the confirmation, `--dry-run`+`-i` note, abort paths, exit codes.
9. Credential retry loop: fail twice, succeed third; abort after 3 failures.

Manual matrix: Linux TTY, `tmux`, Windows Terminal, 80-column terminal, server with
self-signed cert (`IMAP_SSL_VERIFY=false`), tenant with > 10 root folders,
non-ASCII folder name, `Ctrl-C` at every step, migration with 0 matching messages.

## 9. Implementation steps (planned commits)

1. `docs: add interactive TUI spec` (this file) — **review checkpoint**.
2. `feat(tui): add optional --interactive entry point and prompt abstraction`
   (module skeleton, lazy imports, TTY check, `requirements-tui.txt`, no wizard logic yet).
3. `feat(tui): preflight Graph credentials with optional .env write-back`.
4. `feat(tui): mailbox picker, account statistics and scope selection`.
5. `feat(tui): IMAP validation and root-folder picker with fuzzy input`.
6. `feat(tui): validation screen, plan preview and explicit confirmation`.
7. `feat(tui): execute migration and optional CSV row append`.
8. `test(tui): scripted-prompter unit tests`.
9. `docs: document interactive mode and update changelog/TASKS`.

Rough size: ~450–550 lines in `interactive.py`, ~300 lines tests, ~15 lines dispatch.
No changes to batch behaviour.

## 10. Decisions confirmed during review

| # | Decision |
| --- | --- |
| Q1 | Optional `requirements-tui.txt`; batch mode stays dependency-light |
| Q2 | Source mailbox from `details.csv` **and** manual entry |
| Q3 | Persist only the three `SOURCE_*` keys (opt-in); never the IMAP password; ask (default No) about appending the mailbox row to `details.csv` |
| Q4 | **No separate dry-run action in interactive mode.** The final validation screen is the dry run: it verifies credentials, access and the plan without fetching/writing everything. The confirmation always defaults to **No**; only an explicit yes starts the migration. `--dry-run` together with `-i` prints a note and is ignored |
| Q5 | Folder scope as a searchable checkbox list (typed fallback) |
| Q6 | `--folders/--limit/--target-parent-folder` are prefilled, editable defaults |
| Q7 | "Whisper" = typeahead/fuzzy autocomplete in the parent-folder input (not speech) |
| Q8 | 0 messages in scope → report and offer abort, never silently succeed |
