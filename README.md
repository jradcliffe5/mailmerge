# mailmerge

Send personalized email campaigns through Gmail's SMTP service using a CSV data file and text/HTML templates.

## Requirements

- Python 3.8+
- A Gmail account with [2-Step Verification](https://myaccount.google.com/u/0/security) enabled and an [App Password](https://support.google.com/accounts/answer/185833) generated for "Mail".

## Installation

Install directly from the source tree:

```shell
python3 -m pip install .
```

To produce distributable artifacts, run `python3 -m build` and then install the wheel:

```shell
python3 -m pip install dist/gmail_mailmerge-0.1.0-py3-none-any.whl
```

After installation, invoke the CLI with `mailmerge --help` or `python3 -m mailmerge_cli`.

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

3. Run the script:

   ```shell
   python3 mailmerge.py recipients.csv --subject 'Project $project update' --body body.txt --dry-run
   ```

   During a dry run the script prints the full rendered email (headers and body) for inspection. Remove `--dry-run` to send the emails. The script reads the sender address from the `GMAIL_ADDRESS` environment variable and the Gmail app password from `GMAIL_APP_PASSWORD`. You can also pass `--sender` and `--password` explicitly (the password prompt hides your input if you omit `--password`). If you need to use double quotes around the subject, escape dollar signs as `\$` so your shell does not expand them.

### Additional options

- `--body-type html` sends the template as HTML.
- `--delay 1.5` pauses between messages to avoid rate limits.
- `--limit 5` only processes the first 5 rows for testing.
- `--recipient-column contact_email` uses a different column for email addresses.

Run `python3 mailmerge.py --help` to see the full list of flags.
