"""CLI entry point for the Gmail mail merge utility."""

from __future__ import annotations

import argparse
import csv
import getpass
import logging
import os
import smtplib
import sys
import time
from email.message import EmailMessage
from pathlib import Path
from string import Template
from typing import List, Sequence, Tuple


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Send a mail merge to multiple recipients via Gmail SMTP.",
        formatter_class=argparse.RawTextHelpFormatter,
        epilog=(
            "CSV structure:\n"
            "  - Header row defines template variables (required).\n"
            "  - Include a column for recipient addresses (defaults to 'email'; override with --recipient-column).\n"
            "  - Additional columns can be referenced with $column_name placeholders in --subject and the body template.\n"
            "\n"
            "Example:\n"
            "  email,first_name,project\n"
            "  alice@example.com,Alice,Apollo\n"
            "  bob@example.com,Bob,Zephyr\n"
            "\n"
            "Tip: wrap templates that contain $placeholders in single quotes (e.g. --subject 'Hello $name') "
            "or escape dollar signs when your shell performs variable expansion."
        ),
    )
    parser.add_argument(
        "csv",
        type=Path,
        help="CSV file with one row per recipient. Column names become template variables.",
    )
    parser.add_argument(
        "--subject",
        required=True,
        help=(
            "Email subject template. Use Python Template placeholders, e.g. "
            "'Hello $first_name'. Wrap the value in single quotes or escape $ "
            "characters when using shells that expand variables."
        ),
    )
    parser.add_argument(
        "--body",
        type=Path,
        required=True,
        help="Path to a text or HTML template file used as the message body.",
    )
    parser.add_argument(
        "--body-type",
        choices=("plain", "html"),
        default="plain",
        help="Content type for the email body. Defaults to plain text.",
    )
    parser.add_argument(
        "--sender",
        default=os.getenv("GMAIL_ADDRESS"),
        help=(
            "Email address used to send the messages. "
            "Defaults to the GMAIL_ADDRESS environment variable."
        ),
    )
    parser.add_argument(
        "--password",
        default=os.getenv("GMAIL_APP_PASSWORD"),
        help=(
            "App password for the sender account. "
            "Defaults to the GMAIL_APP_PASSWORD environment variable. "
            "Use an app password generated in your Google account settings."
        ),
    )
    parser.add_argument(
        "--recipient-column",
        default="email",
        help="Column name in the CSV that contains recipient email addresses. Defaults to 'email'.",
    )
    parser.add_argument(
        "--reply-to",
        help="Optional Reply-To email address to add to outgoing messages.",
    )
    parser.add_argument(
        "--smtp-server",
        default="smtp.gmail.com",
        help="SMTP server hostname. Defaults to Gmail's SMTP server.",
    )
    parser.add_argument(
        "--smtp-port",
        type=int,
        default=587,
        help="SMTP port number. Defaults to 587 (STARTTLS).",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=0.0,
        help="Optional delay in seconds between messages.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview emails without sending anything.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="Only send to the first N rows in the CSV.",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=("DEBUG", "INFO", "WARNING", "ERROR"),
        help="Logging verbosity. Defaults to INFO.",
    )
    return parser.parse_args(argv)


def load_recipients(csv_path: Path) -> List[dict]:
    if not csv_path.is_file():
        raise FileNotFoundError(f"CSV file not found: {csv_path}")

    with csv_path.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames:
            raise ValueError(f"CSV file {csv_path} is missing a header row.")
        return [row for row in reader]


def read_template(path: Path) -> Template:
    if not path.is_file():
        raise FileNotFoundError(f"Template file not found: {path}")
    content = path.read_text(encoding="utf-8")
    return Template(content)


def parse_addresses(raw: str) -> List[str]:
    return [addr.strip() for addr in raw.replace(";", ",").split(",") if addr.strip()]


def build_message(
    sender: str,
    recipients: Sequence[str],
    subject: str,
    body: str,
    *,
    subtype: str,
    reply_to: str | None = None,
) -> EmailMessage:
    message = EmailMessage()
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    message["Subject"] = subject
    if reply_to:
        message["Reply-To"] = reply_to

    message.set_content(body, subtype=subtype)
    return message


def ensure_credentials(sender: str | None, password: str | None) -> Tuple[str, str]:
    if not sender:
        raise SystemExit(
            "Sender email address is required. Provide --sender or set GMAIL_ADDRESS."
        )
    if not password:
        password = getpass.getpass(
            prompt=f"App password for {sender} (input hidden): ",
        )
    if not password:
        raise SystemExit("Cannot proceed without a password or app password.")
    return sender, password


def format_message(
    template: Template,
    context: dict,
    *,
    template_label: str,
) -> str:
    try:
        return template.substitute(context)
    except KeyError as exc:
        missing = exc.args[0]
        raise ValueError(
            f"Missing '{missing}' for {template_label}. "
            "Ensure the CSV provides this column."
        ) from exc


def send_messages(
    *,
    rows: Sequence[dict],
    sender: str,
    password: str,
    subject_template: Template,
    body_template: Template,
    recipient_column: str,
    smtp_server: str,
    smtp_port: int,
    subtype: str,
    delay: float,
    reply_to: str | None,
    dry_run: bool,
    limit: int | None,
) -> None:
    total = len(rows) if limit is None else min(limit, len(rows))
    if total == 0:
        logging.warning("No recipients found. Nothing to do.")
        return

    logging.info("Preparing to send %s messages.", total)

    if dry_run:
        logging.info("Dry run enabled: emails will be previewed, not sent.")

    if dry_run:
        smtp: smtplib.SMTP | None = None
    else:
        try:
            smtp = smtplib.SMTP(smtp_server, smtp_port)
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(sender, password)
        except Exception:
            logging.exception("Failed to connect or authenticate with the SMTP server.")
            raise

    try:
        for index, row in enumerate(rows[:total], start=1):
            recipient_raw = row.get(recipient_column, "").strip()
            if not recipient_raw:
                logging.warning(
                    "Skipping row %s: missing recipient in column '%s'.",
                    index,
                    recipient_column,
                )
                continue

            recipients = parse_addresses(recipient_raw)
            if not recipients:
                logging.warning(
                    "Skipping row %s: could not parse any addresses from '%s'.",
                    index,
                    recipient_raw,
                )
                continue

            context = {key: value for key, value in row.items() if value is not None}
            context.setdefault("row_number", index)

            subject = format_message(
                subject_template,
                context,
                template_label="subject",
            )
            body = format_message(
                body_template,
                context,
                template_label="body",
            )

            logging.debug("Row %s context: %s", index, context)

            message = build_message(
                sender,
                recipients,
                subject,
                body,
                subtype=subtype,
                reply_to=reply_to,
            )

            if dry_run:
                logging.info(
                    "Dry run %s/%s preview for %s:\n%s",
                    index,
                    total,
                    ", ".join(recipients),
                    message.as_string(),
                )
                continue

            try:
                smtp.send_message(message)
                logging.info(
                    "Sent %s/%s to %s with subject '%s'.",
                    index,
                    total,
                    recipients,
                    subject,
                )
            except Exception:
                logging.exception("Failed to send email for row %s.", index)

            if delay > 0 and index < total:
                logging.debug("Sleeping for %s seconds.", delay)
                time.sleep(delay)
    finally:
        if not dry_run and smtp is not None:
            smtp.quit()


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    logging.basicConfig(level=args.log_level, format="%(levelname)s: %(message)s")

    try:
        rows = load_recipients(args.csv)
        subject_template = Template(args.subject)
        body_template = read_template(args.body)
        if args.dry_run:
            sender = args.sender or "dry-run@example.invalid"
            password = args.password or ""
            if not args.sender:
                logging.debug(
                    "Dry run without sender specified; using placeholder '%s'.",
                    sender,
                )
        else:
            sender, password = ensure_credentials(args.sender, args.password)
    except Exception as exc:
        logging.error("%s", exc)
        return 1

    send_messages(
        rows=rows,
        sender=sender,
        password=password,
        subject_template=subject_template,
        body_template=body_template,
        recipient_column=args.recipient_column,
        smtp_server=args.smtp_server,
        smtp_port=args.smtp_port,
        subtype=args.body_type,
        delay=args.delay,
        reply_to=args.reply_to,
        dry_run=args.dry_run,
        limit=args.limit,
    )
    return 0


__all__ = [
    "main",
    "parse_args",
    "load_recipients",
    "read_template",
    "send_messages",
]
