# mailmerge

Send personalised email campaigns through Gmail's SMTP service using a CSV data file and text/HTML templates.

## Highlights
- Personalise every message with `$variables` pulled from your CSV.
- Send plain text or HTML emails (with automatic HTML detection and wrapping).
- Preview the full email safely with `--dry-run` before delivering it.
- Attach shared or per-recipient files and schedule future sends on cron, launchd, or systemd.

## Requirements
- Python 3.8+
- A Gmail account with [2-Step Verification](https://myaccount.google.com/u/0/security) enabled and an [App Password](https://support.google.com/accounts/answer/185833) generated for "Mail".

## Installation

### Install from the source tree
```shell
python3 -m pip install .
```

### Build and install a wheel (optional)
```shell
python3 -m build
python3 -m pip install dist/gmail_mailmerge-0.9.0-py3-none-any.whl
```

### Run without installing
After cloning the repository you can execute the script directly:
```shell
python3 mailmerge.py --help
```
You can also run the installed console script via `mailmerge --help` or `python3 -m mailmerge_cli`.

## Quick start

### 1. Prepare your recipient data
- Create a CSV file with one row per recipient; column headers become template variables.
- Example (`recipients.csv`):
  ```csv
  email,first_name,project
  alice@example.com,Alice,Apollo
  bob@example.com,Bob,Zephyr
  ```
- Sample CSV and template files live in `examples/` (e.g. `examples/example_contacts.csv`, `examples/example_email.txt`).

### 2. Write your message template
- Use `$variable` placeholders that match the CSV headers.
- Plain text example (`body.txt`):
  ```text
  Hi $first_name,

  Thanks for your work on $project.

  Cheers,
  Jack
  ```
- Files ending in `.html` (or detected as HTML) automatically send as HTML. Plain-text HTML templates are wrapped in a minimal HTML document for consistent rendering. Use `--body-type plain` to force plain text.

### 3. Provide credentials
- By default the CLI reads the sender address from `GMAIL_ADDRESS` and the Gmail app password from `GMAIL_APP_PASSWORD`.
- Override them with `--sender` and `--password` if desired; omitting `--password` triggers a hidden prompt.
- When quoting strings that contain `$`, escape the dollar sign (e.g. `--subject "Project \$project update"`) to prevent shell expansion.

### 4. Run the CLI

#### Installed command
```shell
mailmerge recipients.csv \
  --subject 'Project $project update' \
  --body body.txt \
  --dry-run
```

#### Direct script invocation
```shell
python3 mailmerge.py recipients.csv \
  -s 'Project $project update' \
  -b body.txt \
  -n
```

Removing `--dry-run`/`-n` sends the emails. Dry runs print the full rendered message (headers + body) and use placeholder credentials.

## Command-line reference
Run `python3 mailmerge.py --help` or `mailmerge --help` for the complete usage text. The tables below group the most common flags.

### Core workflow
| Short | Long | Purpose | Notes |
| --- | --- | --- | --- |
| `-s` | `--subject` | Subject template | Uses `$placeholders` from the CSV. |
| `-b` | `--body` | Body template path | Supports plain text and HTML. |
| `-t` | `--body-type` | Body format | Auto-detected; override with `plain` or `html`. |
| `-n` | `--dry-run` | Preview emails without sending | Prints rendered messages to stdout. |

### Recipients and routing
| Short | Long | Purpose | Notes |
| --- | --- | --- | --- |
| `-c` | `--recipient-column` | Choose the recipient address column | Defaults to `email`. |
|  | `--cc` | Additional Cc recipients | Repeatable; accepts comma/semicolon lists and `$placeholders`. |
|  | `--bcc` | Additional Bcc recipients | Repeatable; accepts comma/semicolon lists and `$placeholders`. |
|  | `--cc-column` | Column containing per-recipient Cc values | Values can include comma/semicolon separated addresses. |
|  | `--bcc-column` | Column containing per-recipient Bcc values | Values can include comma/semicolon separated addresses. |
| `-r` | `--reply-to` | Add a Reply-To header | Useful for directing responses elsewhere. |

### Attachments
| Short | Long | Purpose | Notes |
| --- | --- | --- | --- |
| `-a` | `--attachment` | Attach the same file for everyone | Repeat the flag to add multiple files. Paths can include `$placeholders`. |
| `-A` | `--attachment-column` | Attach per-recipient files | Column values can contain comma/semicolon separated paths. Paths resolve relative to the CSV. |

### Credentials and SMTP
| Short | Long | Purpose | Notes |
| --- | --- | --- | --- |
| `-f` | `--sender` | Sender email address | Defaults to `$GMAIL_ADDRESS`. |
| `-p` | `--password` | Gmail app password | Defaults to `$GMAIL_APP_PASSWORD`; prompts if omitted. |
| `-S` | `--smtp-server` | SMTP host | Defaults to `smtp.gmail.com`. |
| `-P` | `--smtp-port` | SMTP port | Defaults to `587` (STARTTLS). |

### Delivery controls
| Short | Long | Purpose | Notes |
| --- | --- | --- | --- |
| `-d` | `--delay` | Delay between emails | Helps avoid rate limits (e.g. `-d 1.5`). |
| `-l` | `--limit` | Process only the first N recipients | Handy for quick tests. |
| `-L` | `--log-level` | Logging verbosity | Accepts `DEBUG`, `INFO`, `WARNING`, `ERROR`. |

### Scheduling
| Short | Long | Purpose | Notes |
| --- | --- | --- | --- |
|  | `--schedule` | Register the command with a scheduler | Accepts cron expressions, ISO times (`09:30`), or ISO datetimes (`2024-06-05T09:30`). |
|  | `--schedule-backend` | Choose scheduler backend | `auto` (default), `cron`, `launchd` (macOS), or `systemd` (Linux). |
|  | `--schedule-timezone` | Interpret ISO schedules in a specific timezone | Provide an IANA tz name (e.g. `Europe/London`). |
|  | `--schedule-label` | Name for the scheduled job | Defaults to the CSV stem; combine with `--schedule-overwrite`. |
|  | `--schedule-overwrite` | Replace an existing job with the same label | Useful for updates. |
|  | `--schedule-list` | Show scheduled entries | Combine with `--schedule-label` to filter results. |
|  | `--schedule-remove` | Delete a specific scheduled job | Pass the original label. |
|  | `--schedule-remove-all` | Remove every mailmerge job for the selected backend | Good for cleanup. |

## Scheduling workflows

Use `--schedule` to run the same command automatically later. The tool auto-detects the best backend (launchd on macOS, systemd on Linux when `systemctl` exists, otherwise cron) and stores lightweight state files to avoid duplicate sends. Plain cron jobs still skip runs while the machine is off.

### Example: weekday cron entry
```shell
mailmerge recipients.csv \
  --subject 'Project $project update' \
  --body body.txt \
  --sender you@example.com \
  --password 'your app password' \
  --schedule '0 9 * * 1-5' \
  --schedule-label project-updates
```
- ISO times (`--schedule 09:30`) run daily at the given local time; ISO datetimes repeat yearly.
- Convert ISO schedules from another timezone with `--schedule-timezone Europe/London`.
- Jobs are stored as a two-line block in `crontab -l`. Re-run with `--schedule-overwrite` to update the command.
- Ensure the cron environment exports `GMAIL_ADDRESS` and `GMAIL_APP_PASSWORD` if you do not hard-code `--sender`/`--password`.

**Python 3.8 note:** install `backports.zoneinfo` or `pytz` if you use `--schedule-timezone`.

### macOS launchd catch-up scheduling
- Set `--schedule-backend launchd` for wake-from-sleep catch-up behaviour.
- Creates `~/Library/LaunchAgents/com.gmailmailmerge.<label>.plist` and logs to `~/Library/Logs/com.gmailmailmerge.<label>.log`.
- Only fixed minute/hour (and optional month/day or weekday) values are supported—ranges like `*/5` are not.
- Launchd retries failed runs automatically with at least 60 seconds between attempts.

Example:
```shell
mailmerge recipients.csv \
  --subject 'Project $project update' \
  --body body.txt \
  --sender you@example.com \
  --password 'your app password' \
  --schedule 09:30 \
  --schedule-backend launchd \
  --schedule-label project-updates
```

### Linux systemd timers with catch-up
- Select `--schedule-backend systemd` to generate user units in `~/.config/systemd/user`.
- Units use `Persistent=true`, so missed runs execute once at the next login/boot.
- Requires `systemctl`; the CLI runs `systemctl --user enable --now` for you.

Example:
```shell
mailmerge recipients.csv \
  --subject 'Project $project update' \
  --body body.txt \
  --sender you@example.com \
  --password 'your app password' \
  --schedule 09:30 \
  --schedule-backend systemd \
  --schedule-label project-updates
```

### Manage scheduled jobs
- List auto-detected jobs: `mailmerge --schedule-list`
- Show launchd jobs: `mailmerge --schedule-backend launchd --schedule-list`
- Remove one job: `mailmerge --schedule-remove project-updates`
- Remove a specific launchd job: `mailmerge --schedule-backend launchd --schedule-remove project-updates`
- Remove everything for the current backend: `mailmerge --schedule-remove-all`
- Clean all launchd jobs: `mailmerge --schedule-backend launchd --schedule-remove-all`
- Clean all systemd jobs: `mailmerge --schedule-backend systemd --schedule-remove-all`

## Examples
Explore `examples/` for a ready-made CSV, plain-text template, and HTML template to adapt for your campaign. They pair well with `--dry-run` while experimenting.
