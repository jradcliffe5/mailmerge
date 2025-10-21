"""CLI entry point for the Gmail mail merge utility."""

from __future__ import annotations

import argparse
import csv
import getpass
import hashlib
import logging
import mimetypes
import os
import shlex
import smtplib
import subprocess
import sys
import time
from datetime import date as dt_date, datetime, time as dt_time, timezone, tzinfo
from email.message import EmailMessage
from pathlib import Path
from string import Template
from typing import List, Sequence, Set, Tuple

try:  # Python 3.9+
    from zoneinfo import ZoneInfo  # type: ignore[attr-defined]
except ImportError:  # pragma: no cover
    try:  # Python 3.8 fallback
        from backports.zoneinfo import ZoneInfo  # type: ignore
    except ImportError:  # pragma: no cover
        ZoneInfo = None  # type: ignore

try:  # Optional pure-Python fallback
    import pytz  # type: ignore
except ImportError:  # pragma: no cover
    pytz = None  # type: ignore


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
        nargs="?",
        help="CSV file with one row per recipient. Column names become template variables.",
    )
    parser.add_argument(
        "-s",
        "--subject",
        help=(
            "Email subject template. Use Python Template placeholders, e.g. "
            "'Hello $first_name'. Wrap the value in single quotes or escape $ "
            "characters when using shells that expand variables."
        ),
    )
    parser.add_argument(
        "-b",
        "--body",
        type=Path,
        help="Path to a text or HTML template file used as the message body.",
    )
    parser.add_argument(
        "-t",
        "--body-type",
        choices=("plain", "html"),
        default="plain",
        help="Content type for the email body. Defaults to plain text.",
    )
    parser.add_argument(
        "-a",
        "--attachment",
        dest="attachments",
        action="append",
        default=[],
        help=(
            "Path template for an attachment. Can be provided multiple times. "
            "Supports $placeholders from the CSV."
        ),
    )
    parser.add_argument(
        "-A",
        "--attachment-column",
        help=(
            "CSV column that lists attachment paths for each recipient. "
            "Separate multiple paths with commas or semicolons."
        ),
    )
    parser.add_argument(
        "-f",
        "--sender",
        default=os.getenv("GMAIL_ADDRESS"),
        help=(
            "Email address used to send the messages. "
            "Defaults to the GMAIL_ADDRESS environment variable."
        ),
    )
    parser.add_argument(
        "-p",
        "--password",
        default=os.getenv("GMAIL_APP_PASSWORD"),
        help=(
            "App password for the sender account. "
            "Defaults to the GMAIL_APP_PASSWORD environment variable. "
            "Use an app password generated in your Google account settings."
        ),
    )
    parser.add_argument(
        "-c",
        "--recipient-column",
        default="email",
        help="Column name in the CSV that contains recipient email addresses. Defaults to 'email'.",
    )
    parser.add_argument(
        "-r",
        "--reply-to",
        help="Optional Reply-To email address to add to outgoing messages.",
    )
    parser.add_argument(
        "-S",
        "--smtp-server",
        default="smtp.gmail.com",
        help="SMTP server hostname. Defaults to Gmail's SMTP server.",
    )
    parser.add_argument(
        "-P",
        "--smtp-port",
        type=int,
        default=587,
        help="SMTP port number. Defaults to 587 (STARTTLS).",
    )
    parser.add_argument(
        "-d",
        "--delay",
        type=float,
        default=0.0,
        help="Optional delay in seconds between messages.",
    )
    parser.add_argument(
        "-n",
        "--dry-run",
        action="store_true",
        help="Preview emails without sending anything.",
    )
    parser.add_argument(
        "-l",
        "--limit",
        type=int,
        help="Only send to the first N rows in the CSV.",
    )
    parser.add_argument(
        "-L",
        "--log-level",
        default="INFO",
        choices=("DEBUG", "INFO", "WARNING", "ERROR"),
        help="Logging verbosity. Defaults to INFO.",
    )
    parser.add_argument(
        "--schedule",
        metavar="CRON",
        help=(
            "Install this invocation in the user's crontab instead of sending immediately. "
            "Provide a standard 5-field cron expression (e.g. '0 9 * * 1-5'), an @-macro, "
            "or an ISO 8601 time/datetime such as '09:30' or '2024-06-05T09:30'."
        ),
    )
    parser.add_argument(
        "--schedule-timezone",
        help=(
            "Timezone identifier (IANA database) used when interpreting ISO schedule values, "
            "e.g. 'Europe/London'. Defaults to the system timezone."
        ),
    )
    parser.add_argument(
        "--schedule-label",
        help=(
            "Identifier used to mark the cron entry. Defaults to the CSV file stem. "
            "Combine with --schedule-overwrite to update an existing entry."
        ),
    )
    parser.add_argument(
        "--schedule-overwrite",
        action="store_true",
        help="Replace an existing cron entry that uses the same schedule label.",
    )
    parser.add_argument(
        "--schedule-remove-all",
        action="store_true",
        help="Remove every cron entry created by mailmerge. Other cron jobs remain untouched.",
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


def parse_list_entries(raw: str | None) -> List[str]:
    if not raw:
        return []
    return [value.strip() for value in raw.replace(";", ",").split(",") if value.strip()]


def build_message(
    sender: str,
    recipients: Sequence[str],
    subject: str,
    body: str,
    *,
    subtype: str,
    reply_to: str | None = None,
    attachments: Sequence[Tuple[str, bytes, str, str]] | None = None,
) -> EmailMessage:
    message = EmailMessage()
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    message["Subject"] = subject
    if reply_to:
        message["Reply-To"] = reply_to

    message.set_content(body, subtype=subtype)
    if attachments:
        for filename, content, maintype, subtype_name in attachments:
            message.add_attachment(
                content,
                maintype=maintype,
                subtype=subtype_name,
                filename=filename,
            )
    return message


CRON_COMMENT_PREFIX = "# mailmerge schedule:"


def build_cron_command(args: argparse.Namespace) -> str:
    parts: List[str] = [
        shlex.quote(sys.executable),
        "-m",
        "mailmerge_cli",
        shlex.quote(str(args.csv)),
        "--subject",
        shlex.quote(args.subject),
        "--body",
        shlex.quote(str(args.body)),
    ]

    if args.body_type != "plain":
        parts.extend(["--body-type", shlex.quote(args.body_type)])

    for attachment in args.attachments:
        parts.extend(["--attachment", shlex.quote(attachment)])

    if args.attachment_column:
        parts.extend(["--attachment-column", shlex.quote(args.attachment_column)])

    if args.sender:
        parts.extend(["--sender", shlex.quote(args.sender)])

    if args.password:
        parts.extend(["--password", shlex.quote(args.password)])

    if args.recipient_column != "email":
        parts.extend(["--recipient-column", shlex.quote(args.recipient_column)])

    if args.reply_to:
        parts.extend(["--reply-to", shlex.quote(args.reply_to)])

    if args.smtp_server != "smtp.gmail.com":
        parts.extend(["--smtp-server", shlex.quote(args.smtp_server)])

    if args.smtp_port != 587:
        parts.extend(["--smtp-port", shlex.quote(str(args.smtp_port))])

    if args.delay > 0:
        parts.extend(["--delay", shlex.quote(str(args.delay))])

    if args.dry_run:
        parts.append("--dry-run")

    if args.limit is not None:
        parts.extend(["--limit", shlex.quote(str(args.limit))])

    if args.log_level.upper() != "INFO":
        parts.extend(["--log-level", shlex.quote(args.log_level.upper())])

    return " ".join(parts)


def read_crontab_lines() -> List[str]:
    result = subprocess.run(
        ["crontab", "-l"],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode == 0:
        return result.stdout.splitlines()
    stderr = result.stderr.strip().lower()
    if result.returncode == 1 and "no crontab" in stderr:
        return []
    raise RuntimeError(f"Unable to read crontab: {result.stderr.strip() or result.stdout.strip()}")


def write_crontab_lines(lines: Sequence[str]) -> None:
    new_content = "\n".join(lines).rstrip("\n") + "\n"
    result = subprocess.run(
        ["crontab", "-"],
        input=new_content,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"Failed to update crontab: {result.stderr.strip() or result.stdout.strip()}")


def resolve_schedule_timezone(name: str | None) -> tzinfo | None:
    if not name:
        return None
    if ZoneInfo is None:
        if pytz is None:
            raise RuntimeError(
                "Timezone support requires Python 3.9+ (zoneinfo module), the 'backports.zoneinfo' package, "
                "or 'pytz'. Install one of these to use --schedule-timezone."
            )
        try:
            zone = pytz.timezone(name)
        except Exception as exc:
            raise ValueError(f"Unknown timezone identifier '{name}'.") from exc
        setattr(zone, "key", getattr(zone, "key", getattr(zone, "zone", name)))
        return zone
    try:
        return ZoneInfo(name)
    except Exception:
        if pytz is None:
            raise
        try:
            zone = pytz.timezone(name)
        except Exception as exc:
            raise ValueError(f"Unknown timezone identifier '{name}'.") from exc
        setattr(zone, "key", getattr(zone, "key", getattr(zone, "zone", name)))
        return zone


def describe_timezone(zone: tzinfo | None) -> str:
    if zone is None:
        return "UTC"
    key = getattr(zone, "key", None)
    if key:
        return str(key)
    zone_name = getattr(zone, "zone", None)
    if zone_name:
        return str(zone_name)
    name = zone.tzname(None)
    if name:
        return name
    return str(zone)


def ensure_timezone(zone: tzinfo | None) -> tzinfo:
    if zone is None:
        return timezone.utc
    return zone


def ensure_datetime_timezone(dt_value: datetime, zone: tzinfo | None) -> datetime:
    if dt_value.tzinfo is not None:
        if zone is None:
            return dt_value
        return dt_value.astimezone(ensure_timezone(zone))
    tz_obj = ensure_timezone(zone)
    localize = getattr(tz_obj, "localize", None)
    if callable(localize):
        return localize(dt_value)  # type: ignore[call-arg]
    return dt_value.replace(tzinfo=tz_obj)


def combine_time_with_timezone(date_value: dt_date, time_value: dt_time, zone: tzinfo | None) -> datetime:
    if time_value.tzinfo is not None:
        naive_time = time_value.replace(tzinfo=None)
        base_dt = datetime.combine(date_value, naive_time)
        return ensure_datetime_timezone(base_dt, time_value.tzinfo)
    base_dt = datetime.combine(date_value, time_value)
    return ensure_datetime_timezone(base_dt, zone)


def normalize_schedule(schedule: str, tz_hint: tzinfo | None, local_zone: tzinfo) -> tuple[str, str]:
    cleaned = schedule.strip()
    if not cleaned or any(char in cleaned for char in "\r\n"):
        raise ValueError("Schedule expression must be a single non-empty line.")

    if cleaned.startswith("@"):
        return cleaned, cleaned

    fields = cleaned.split()
    if len(fields) in {5, 6}:
        return cleaned, cleaned

    cron_from_iso = convert_iso_to_cron(cleaned, tz_hint, local_zone)
    if cron_from_iso is None:
        raise ValueError(
            "Schedule must be a cron expression (five fields, or @daily/@hourly etc.) "
            "or an ISO 8601 time/datetime such as '09:30' or '2024-06-05T09:30'."
        )
    return cron_from_iso, cleaned


def convert_iso_to_cron(value: str, tz_hint: tzinfo | None, local_zone: tzinfo) -> str | None:
    candidate = value.strip()
    if candidate.endswith("Z") and candidate.count("Z") == 1:
        candidate = candidate[:-1] + "+00:00"

    try:
        dt_value = datetime.fromisoformat(candidate)
    except ValueError:
        dt_value = None

    if dt_value is not None:
        source_zone = dt_value.tzinfo or tz_hint or local_zone
        dt_with_tz = ensure_datetime_timezone(dt_value, source_zone)
        dt_local = dt_with_tz.astimezone(local_zone)
        return f"{dt_local.minute} {dt_local.hour} {dt_local.day} {dt_local.month} *"

    time_candidate = value.strip()
    if time_candidate.startswith("T"):
        time_candidate = time_candidate[1:]
    if time_candidate.endswith("Z") and time_candidate.count("Z") == 1:
        time_candidate = time_candidate[:-1] + "+00:00"

    try:
        time_value = dt_time.fromisoformat(time_candidate)
    except ValueError:
        return None

    source_zone = time_value.tzinfo or tz_hint or local_zone
    reference_date = datetime.now(ensure_timezone(source_zone)).date()
    combined = combine_time_with_timezone(reference_date, time_value, source_zone)
    local_time = combined.astimezone(local_zone)

    return f"{local_time.minute} {local_time.hour} * * *"


def install_cron_job(args: argparse.Namespace) -> None:
    schedule = args.schedule.strip()

    local_zone = datetime.now().astimezone().tzinfo or timezone.utc
    tz_hint = resolve_schedule_timezone(args.schedule_timezone)

    cron_spec, display_schedule = normalize_schedule(schedule, tz_hint, local_zone)

    command = build_cron_command(args)
    label_seed = args.schedule_label or args.csv.stem or "mailmerge"
    fingerprint = hashlib.sha1(command.encode("utf-8")).hexdigest()[:8]
    comment_label = f"{CRON_COMMENT_PREFIX} {label_seed}"
    comment_line = f"{comment_label} ({fingerprint})"

    lines = read_crontab_lines()
    existing_index = next(
        (index for index, line in enumerate(lines) if line.startswith(comment_label)),
        None,
    )

    if existing_index is not None:
        if not args.schedule_overwrite:
            raise RuntimeError(
                f"Cron entry already exists for label '{label_seed}'. "
                "Re-run with --schedule-overwrite to replace it."
            )
        del lines[existing_index]
        if existing_index < len(lines):
            del lines[existing_index]
        while existing_index < len(lines) and not lines[existing_index].strip():
            del lines[existing_index]

    if lines and lines[-1].strip():
        lines.append("")
    lines.extend([comment_line, f"{cron_spec} {command}"])
    write_crontab_lines(lines)

    system_zone_label = describe_timezone(local_zone)
    if args.schedule_timezone:
        logging.info(
            "Installed cron entry '%s' (%s interpreted in %s) using cron spec '%s' (system timezone %s).",
            label_seed,
            display_schedule,
            args.schedule_timezone,
            cron_spec,
            system_zone_label,
        )
    else:
        logging.info(
            "Installed cron entry '%s' (%s) using cron spec '%s' (system timezone %s).",
            label_seed,
            display_schedule,
            cron_spec,
            system_zone_label,
        )
    if not args.password and not os.getenv("GMAIL_APP_PASSWORD"):
        logging.warning(
            "Cron job does not include a password. Ensure the job environment defines GMAIL_APP_PASSWORD."
        )
    if not args.sender and not os.getenv("GMAIL_ADDRESS"):
        logging.warning(
            "Cron job does not include a sender address. Ensure the job environment defines GMAIL_ADDRESS."
        )


def remove_mailmerge_cron_jobs() -> int:
    lines = read_crontab_lines()
    if not lines:
        return 0

    new_lines: List[str] = []
    skip_next = False
    removed_jobs = 0

    for line in lines:
        if skip_next:
            skip_next = False
            continue
        if line.startswith(CRON_COMMENT_PREFIX):
            skip_next = True
            removed_jobs += 1
            if new_lines and not new_lines[-1].strip():
                new_lines.pop()
            continue
        new_lines.append(line)

    while new_lines and not new_lines[-1].strip():
        new_lines.pop()

    if removed_jobs > 0:
        write_crontab_lines(new_lines)

    return removed_jobs


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
    attachment_templates: Sequence[Template],
    attachment_column: str | None,
    attachment_base: Path,
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

            attachment_values: List[str] = []
            attachment_payloads: List[Tuple[str, bytes, str, str]]

            attachments_valid = True
            for template in attachment_templates:
                try:
                    rendered = format_message(
                        template,
                        context,
                        template_label="attachment",
                    )
                except ValueError as exc:
                    logging.error("Skipping row %s: %s", index, exc)
                    attachments_valid = False
                    break
                attachment_values.extend(parse_list_entries(rendered))

            if not attachments_valid:
                continue

            if attachment_column:
                attachment_values.extend(parse_list_entries(row.get(attachment_column)))

            resolved_paths: List[Path] = []
            seen_paths: Set[Path] = set()
            for raw_path in attachment_values:
                candidate = Path(raw_path).expanduser()
                if not candidate.is_absolute():
                    candidate = (attachment_base / candidate).resolve()
                else:
                    candidate = candidate.resolve()
                if candidate in seen_paths:
                    continue
                seen_paths.add(candidate)
                resolved_paths.append(candidate)

            attachment_payloads = []
            for attachment_path in resolved_paths:
                if not attachment_path.exists():
                    logging.error(
                        "Skipping row %s: attachment not found: %s",
                        index,
                        attachment_path,
                    )
                    attachments_valid = False
                    break
                if not attachment_path.is_file():
                    logging.error(
                        "Skipping row %s: attachment path is not a file: %s",
                        index,
                        attachment_path,
                    )
                    attachments_valid = False
                    break
                try:
                    data = attachment_path.read_bytes()
                except OSError as exc:
                    logging.error(
                        "Skipping row %s: unable to read attachment %s (%s)",
                        index,
                        attachment_path,
                        exc,
                    )
                    attachments_valid = False
                    break
                mime_type, _ = mimetypes.guess_type(str(attachment_path))
                if mime_type:
                    maintype, subtype = mime_type.split("/", 1)
                else:
                    maintype, subtype = "application", "octet-stream"
                attachment_payloads.append(
                    (attachment_path.name, data, maintype, subtype)
                )

            if not attachments_valid:
                continue

            message = build_message(
                sender,
                recipients,
                subject,
                body,
                subtype=subtype,
                reply_to=reply_to,
                attachments=attachment_payloads,
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

    schedule_requested = bool(args.schedule)
    remove_requested = bool(args.schedule_remove_all)

    if remove_requested and schedule_requested:
        logging.error("Cannot combine --schedule with --schedule-remove-all.")
        return 1

    if remove_requested and (args.schedule_label or args.schedule_overwrite):
        logging.error("--schedule-remove-all cannot be combined with other scheduling flags.")
        return 1

    if remove_requested:
        try:
            removed = remove_mailmerge_cron_jobs()
        except Exception as exc:
            logging.error("%s", exc)
            return 1
        if removed:
            plural = "entry" if removed == 1 else "entries"
            logging.info("Removed %s scheduled mailmerge cron %s.", removed, plural)
        else:
            logging.info("No mailmerge cron entries found to remove.")
        return 0

    if args.csv is None:
        logging.error("CSV file is required unless --schedule-remove-all is provided.")
        return 1
    if args.subject is None:
        logging.error("Subject template is required.")
        return 1
    if args.body is None:
        logging.error("Body template path is required.")
        return 1

    args.csv = args.csv.expanduser().resolve()
    args.body = args.body.expanduser().resolve()

    if not args.csv.is_file():
        logging.error("CSV file not found: %s", args.csv)
        return 1
    if not args.body.is_file():
        logging.error("Body template file not found: %s", args.body)
        return 1

    if args.schedule:
        try:
            install_cron_job(args)
        except Exception as exc:
            logging.error("%s", exc)
            return 1
        return 0

    try:
        rows = load_recipients(args.csv)
        subject_template = Template(args.subject)
        body_template = read_template(args.body)
        attachment_templates = [Template(value) for value in args.attachments]
        attachment_base = args.csv.parent.resolve()
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
        attachment_templates=attachment_templates,
        attachment_column=args.attachment_column,
        attachment_base=attachment_base,
    )
    return 0


__all__ = [
    "main",
    "parse_args",
    "load_recipients",
    "read_template",
    "send_messages",
]
