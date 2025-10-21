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
python3 -m pip install dist/gmail_mailmerge-0.3.0-py3-none-any.whl
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
| `-r` | `--reply-to` | Reply-To header | Adds a Reply-To address to outgoing emails. |
| `-S` | `--smtp-server` | SMTP host | Defaults to `smtp.gmail.com`. |
| `-P` | `--smtp-port` | SMTP port | Defaults to `587` (STARTTLS). |
| `-d` | `--delay` | Delay between emails | Useful to avoid rate limits (e.g. `-d 1.5`). |
| `-l` | `--limit` | Limit number of recipients processed | Helpful for quick tests. |
| `-L` | `--log-level` | Logging verbosity | Accepts `DEBUG`, `INFO`, `WARNING`, `ERROR`. |
| `-n` | `--dry-run` | Preview emails without sending | Prints rendered messages to stdout. |

Run `python3 mailmerge.py --help` to see the full list of flags.
