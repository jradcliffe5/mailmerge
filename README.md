# mailmerge

Send personalised email campaigns through Gmail's SMTP service using a CSV data file and text/HTML templates.

## Requirements

- Python 3.8+
- A Gmail account with [2-Step Verification](https://myaccount.google.com/u/0/security) enabled and an [App Password](https://support.google.com/accounts/answer/185833) generated for "Mail".

## Installation

Install directly from the source tree:

```shell
python3 -m pip install .
```

To produce distributable artefacts, run `python3 -m build` and then install the wheel:

```shell
python3 -m pip install dist/gmail_mailmerge-0.5.0-py3-none-any.whl
```

After installation, invoke the CLI with `mailmerge --help` or `python3 -m mailmerge_cli`. If you prefer not to install the package, you can run the bundled script directly with `python3 mailmerge.py`.

## Usage

1. Create a CSV file with one recipient per row. The column headers become template variables:

   ```csv
   email,first_name,project
   alice@example.com,Alice,Apollo
   bob@example.com,Bob,Zephyr
   ```

2. Create a message body template (plain text or HTML) using `$variable` placeholders. Example `body.txt`:

   ```text
   Hi $first_name,

   Thanks for your work on $project.

   Cheers,
   Jack
   ```

3. Choose how to run the tool:

   - Installed CLI via pip:

     ```shell
     mailmerge recipients.csv --subject 'Project $project update' --body body.txt --dry-run
     ```

   - Direct script invocation:

     ```shell
     python3 mailmerge.py recipients.csv -s 'Project $project update' -b body.txt -n
     ```

   During a dry run the command prints the full rendered email (headers and body) for inspection. When `--dry-run` is used, you can omit `--sender` and `--password`; a placeholder sender address is used purely for previewing. Remove `--dry-run` to send the emails. Both entry points read the sender address from the `GMAIL_ADDRESS` environment variable and the Gmail app password from `GMAIL_APP_PASSWORD`. You can also pass `--sender` and `--password` explicitly (the password prompt hides your input if you omit `--password`). If you need to use double quotes around the subject, escape dollar signs as `\$` so your shell does not expand them.

### Additional options

| Short | Long | Purpose | Notes |
| --- | --- | --- | --- |
| `-s` | `--subject` | Subject template | Uses `$placeholders` from the CSV. |
| `-b` | `--body` | Body template path | Reads either plain text or HTML. |
| `-t` | `--body-type` | Body format | Use `html` to send HTML emails (default is plain text). |
| `-a` | `--attachment` | Static attachment path(s) | Repeat the flag to add more files. Paths can include `$placeholders` and are resolved relative to the CSV file. |
| `-A` | `--attachment-column` | Per-recipient attachments | Column can contain comma/semicolon delimited paths. |
| `-f` | `--sender` | Sender email address | Defaults to `$GMAIL_ADDRESS`. |
| `-p` | `--password` | App password | Defaults to `$GMAIL_APP_PASSWORD`; prompts if omitted. |
| `-c` | `--recipient-column` | Recipient address column | Falls back to `email`. |
|  | `--cc` | Additional Cc recipients | Repeatable; accepts comma/semicolon lists and supports `$placeholders`. |
|  | `--bcc` | Additional Bcc recipients | Repeatable; accepts comma/semicolon lists and supports `$placeholders`. |
|  | `--cc-column` | Per-recipient Cc column | Column values can contain comma/semicolon separated addresses. |
|  | `--bcc-column` | Per-recipient Bcc column | Column values can contain comma/semicolon separated addresses. |
| `-r` | `--reply-to` | Reply-To header | Adds a Reply-To address to outgoing emails. |
| `-S` | `--smtp-server` | SMTP host | Defaults to `smtp.gmail.com`. |
| `-P` | `--smtp-port` | SMTP port | Defaults to `587` (STARTTLS). |
| `-d` | `--delay` | Delay between emails | Useful to avoid rate limits (e.g. `-d 1.5`). |
| `-l` | `--limit` | Limit number of recipients processed | Helpful for quick tests. |
| `-L` | `--log-level` | Logging verbosity | Accepts `DEBUG`, `INFO`, `WARNING`, `ERROR`. |
| `-n` | `--dry-run` | Preview emails without sending | Prints rendered messages to stdout. |
|  | `--schedule` | Install command in crontab | Provide a cron expression (`0 9 * * 1-5`), an `@daily`/`@hourly` macro, or an ISO time like `09:30` / `2024-06-05T09:30`. |
|  | `--schedule-backend` | Scheduler backend | `auto` (default), `cron`, `launchd` (macOS), or `systemd` (Linux). |
|  | `--schedule-timezone` | Timezone for ISO schedules | IANA tz name (e.g. `Europe/London`) used when parsing ISO times/dates. |
|  | `--schedule-label` | Cron entry identifier | Defaults to the CSV file stem; combine with `--schedule-overwrite`. |
|  | `--schedule-overwrite` | Replace existing cron entry | Updates the cron job that matches `--schedule-label`. |
|  | `--schedule-remove-all` | Delete mailmerge entries | Removes every scheduled entry created by this tool for the selected backend. |

Run `python3 mailmerge.py --help` to see the full list of flags.

### Scheduling with cron, launchd, or systemd

Use `--schedule` to register the current invocation so the emails deliver later from the same machine. By default the tool auto-detects the scheduler (launchd on macOS, systemd on Linux when `systemctl` is available, otherwise cron). Pass `--schedule-backend cron` (or `launchd` / `systemd`) to override this behaviour explicitly. A lightweight state file ensures each job runs at most once per scheduled slot and enables catch-up for launchd/systemd when the original time was missed (plain cron still skips runs while the machine is off):

```shell
mailmerge recipients.csv \
  --subject 'Project $project update' \
  --body body.txt \
  --sender you@example.com \
  --password 'your app password' \
  --schedule "0 9 * * 1-5" \
  --schedule-label project-updates
```

Instead of a cron expression you can also use ISO 8601 times (`--schedule 09:30` to send daily at 09:30) or datetimes (`--schedule 2024-06-05T09:30`). Use `--schedule-timezone` when the ISO value should be interpreted in a specific timezone—mailmerge converts it to the machine’s local timezone before writing the cron entry (e.g. `--schedule 09:30 --schedule-timezone Europe/London`). For ISO datetimes, cron repeats the job yearly on the same calendar date—delete the entry after it runs if you only need a one-off. The command stores a two-line block in `crontab -l`: a marker comment and the full mailmerge command. Re-running with the same `--schedule-label` and `--schedule-overwrite` refreshes the job. If you omit `--sender` or `--password`, ensure the cron environment exports `GMAIL_ADDRESS` and `GMAIL_APP_PASSWORD` before the job runs. The machine must stay powered, logged in, and connected to the internet at the scheduled time. Remove the job later with `mailmerge --schedule-remove-all` (or via `crontab -e` to delete the comment/command pair manually).

> **Note:** Python 3.8 users need either `backports.zoneinfo` or `pytz` to resolve IANA timezones (e.g. `python3 -m pip install backports.zoneinfo` or `python3 -m pip install pytz`).

#### macOS launchd catch-up scheduling

On macOS you can switch to `launchd`, which retries soon after the laptop wakes up if it was asleep at the scheduled time:

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

The tool installs `~/Library/LaunchAgents/com.gmailmailmerge.<label>.plist` and logs to `~/Library/Logs/com.gmailmailmerge.<label>.log`. launchd only accepts fixed minute/hour (and optional month/day or weekday) values—ranges such as `*/5` are not supported here. If the Mac was asleep or powered off when the scheduled time passed, the job runs once shortly after the machine wakes thanks to the stored schedule state. Remove the job later with `mailmerge --schedule-backend launchd --schedule-remove-all` (or provide a different label and overwrite it).

#### Linux systemd timers with catch-up

On Linux you can target systemd, which writes user units under `~/.config/systemd/user` and enables `Persistent=true` so missed runs fire immediately at boot/login:

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

This creates `mailmerge-project-updates.service` and `.timer`, runs `systemctl --user enable --now`, and expects `systemctl` to be available. As with launchd, supply explicit minute/hour values (and optionally month/day *or* weekday). The timer is configured with `Persistent=true`, so a missed run fires once at the next login/boot. Check the timer with `systemctl --user status mailmerge-project-updates.timer` and remove all mailmerge timers with `mailmerge --schedule-backend systemd --schedule-remove-all` (or rely on the default `auto`, which will choose systemd on Linux).

To delete every scheduled entry previously added by mailmerge without touching any other jobs, run:

```shell
mailmerge --schedule-remove-all                                  # uses auto-detected backend
mailmerge --schedule-backend launchd --schedule-remove-all       # macOS launchd
mailmerge --schedule-backend systemd --schedule-remove-all       # Linux systemd
```
