"""CLI entry point for the Gmail mail merge utility."""

from __future__ import annotations

import argparse
import csv
import getpass
import hashlib
import json
import logging
import mimetypes
import os
import plistlib
import re
import shlex
import shutil
import smtplib
import subprocess
import sys
import time
from datetime import date as dt_date, datetime, time as dt_time, timedelta, timezone, tzinfo
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
        "--cc",
        dest="cc",
        action="append",
        default=[],
        help=(
            "Additional Cc recipient address template. Can be provided multiple times. "
            "Supports comma/semicolon separated lists and $placeholders."
        ),
    )
    parser.add_argument(
        "--bcc",
        dest="bcc",
        action="append",
        default=[],
        help=(
            "Additional Bcc recipient address template. Can be provided multiple times. "
            "Supports comma/semicolon separated lists and $placeholders."
        ),
    )
    parser.add_argument(
        "--cc-column",
        help=(
            "CSV column that lists Cc addresses for each recipient. "
            "Separate multiple addresses with commas or semicolons."
        ),
    )
    parser.add_argument(
        "--bcc-column",
        help=(
            "CSV column that lists Bcc addresses for each recipient. "
            "Separate multiple addresses with commas or semicolons."
        ),
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
        "--schedule-backend",
        choices=("auto", "cron", "launchd", "systemd"),
        default="auto",
        help=(
            "Scheduler to configure when using --schedule. "
            "Defaults to 'auto' (launchd on macOS, systemd on Linux if available, otherwise cron). "
            "Use 'cron', 'launchd', or 'systemd' to force a specific backend."
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
        help=(
            "Remove every scheduled entry created by mailmerge for the selected backend "
            "(backend defaults to auto-detected unless overridden)."
        ),
    )
    parser.add_argument(
        "--schedule-spec",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--schedule-state",
        type=Path,
        help=argparse.SUPPRESS,
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
    cc: Sequence[str] | None = None,
    bcc: Sequence[str] | None = None,
    attachments: Sequence[Tuple[str, bytes, str, str]] | None = None,
) -> EmailMessage:
    message = EmailMessage()
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    message["Subject"] = subject
    if reply_to:
        message["Reply-To"] = reply_to
    if cc:
        message["Cc"] = ", ".join(cc)
    if bcc:
        message["Bcc"] = ", ".join(bcc)

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
LAUNCHD_LABEL_PREFIX = "com.gmailmailmerge."
SYSTEMD_UNIT_PREFIX = "mailmerge-"


def sanitize_label_seed(value: str) -> str:
    sanitized = re.sub(r"[^A-Za-z0-9_.-]+", "-", value)
    sanitized = sanitized.strip(".-")
    return sanitized or "mailmerge"


def build_program_arguments(args: argparse.Namespace, *, schedule_spec: str | None = None, state_path: Path | None = None) -> List[str]:
    command: List[str] = [
        sys.executable,
        "-m",
        "mailmerge_cli",
        str(args.csv),
        "--subject",
        args.subject,
        "--body",
        str(args.body),
    ]

    if args.body_type != "plain":
        command.extend(["--body-type", args.body_type])

    for attachment in args.attachments:
        command.extend(["--attachment", str(attachment)])

    if args.attachment_column:
        command.extend(["--attachment-column", str(args.attachment_column)])

    if args.sender:
        command.extend(["--sender", str(args.sender)])

    if args.password:
        command.extend(["--password", str(args.password)])

    if args.recipient_column != "email":
        command.extend(["--recipient-column", str(args.recipient_column)])

    for value in getattr(args, "cc", []):
        command.extend(["--cc", str(value)])

    for value in getattr(args, "bcc", []):
        command.extend(["--bcc", str(value)])

    if args.cc_column:
        command.extend(["--cc-column", str(args.cc_column)])

    if args.bcc_column:
        command.extend(["--bcc-column", str(args.bcc_column)])

    if args.reply_to:
        command.extend(["--reply-to", str(args.reply_to)])

    if args.smtp_server != "smtp.gmail.com":
        command.extend(["--smtp-server", str(args.smtp_server)])

    if args.smtp_port != 587:
        command.extend(["--smtp-port", str(args.smtp_port)])

    if args.delay > 0:
        command.extend(["--delay", str(args.delay)])

    if args.dry_run:
        command.append("--dry-run")

    if args.limit is not None:
        command.extend(["--limit", str(args.limit)])

    if args.log_level.upper() != "INFO":
        command.extend(["--log-level", args.log_level.upper()])

    if schedule_spec:
        command.extend(["--schedule-spec", schedule_spec])
    if state_path:
        command.extend(["--schedule-state", str(state_path)])

    return command


def build_cron_command(args: argparse.Namespace, *, schedule_spec: str | None = None, state_path: Path | None = None) -> str:
    parts = build_program_arguments(args, schedule_spec=schedule_spec, state_path=state_path)
    return " ".join(shlex.quote(part) for part in parts)


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


def require_backend_supported(backend: str) -> None:
    if backend == "launchd" and sys.platform != "darwin":
        raise RuntimeError("The launchd backend is only available on macOS.")
    if backend == "systemd" and not sys.platform.startswith("linux"):
        raise RuntimeError("The systemd backend is only available on Linux.")


def detect_default_backend() -> str:
    if sys.platform == "darwin":
        return "launchd"
    if sys.platform.startswith("linux") and shutil.which("systemctl"):
        return "systemd"
    return "cron"


def determine_backend(choice: str | None) -> str:
    if not choice or choice == "auto":
        backend = detect_default_backend()
    else:
        backend = choice
    require_backend_supported(backend)
    return backend


def schedule_state_base_dir() -> Path:
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "mailmerge"
    base_dir = Path.home() / ".local" / "share" / "mailmerge"
    return base_dir


def schedule_state_dir() -> Path:
    path = schedule_state_base_dir() / "schedule-state"
    path.mkdir(parents=True, exist_ok=True)
    return path


def build_state_path(label_seed: str, fingerprint: str) -> Path:
    filename = f"{label_seed}-{fingerprint}.json"
    return schedule_state_dir() / filename


def parse_cron_value(value: str) -> int | None:
    if value == "*":
        return None
    parsed = int(value)
    return parsed


def cron_matches(dt_value: datetime, minute: int | None, hour: int | None, day: int | None, month: int | None, weekday: int | None) -> bool:
    if minute is not None and dt_value.minute != minute:
        return False
    if hour is not None and dt_value.hour != hour:
        return False
    if month is not None and dt_value.month != month:
        return False
    if day is not None and dt_value.day != day:
        return False
    if weekday is not None:
        cron_weekday = (dt_value.weekday() + 1) % 7
        target = weekday % 7
        if cron_weekday != target:
            return False
    return True


def parse_cron_spec(cron_spec: str) -> tuple[int | None, int | None, int | None, int | None, int | None]:
    minute_s, hour_s, day_s, month_s, weekday_s = cron_spec.split()
    return (
        parse_cron_value(minute_s),
        parse_cron_value(hour_s),
        parse_cron_value(day_s),
        parse_cron_value(month_s),
        parse_cron_value(weekday_s),
    )


def next_run_time(cron_spec: str, start: datetime) -> datetime:
    minute, hour, day, month, weekday = parse_cron_spec(cron_spec)
    candidate = (start + timedelta(minutes=1)).replace(second=0, microsecond=0)
    limit = candidate + timedelta(days=366)
    while candidate <= limit:
        if cron_matches(candidate, minute, hour, day, month, weekday):
            return candidate
        candidate += timedelta(minutes=1)
    raise RuntimeError(f"Unable to compute next run time for cron spec: {cron_spec}")


def initialize_schedule_state(state_path: Path, cron_spec: str, *, overwrite: bool) -> datetime:
    state_path.parent.mkdir(parents=True, exist_ok=True)
    if state_path.exists() and not overwrite:
        try:
            with state_path.open("r", encoding="utf-8") as handle:
                data = json.load(handle)
            next_due_str = data.get("next_due")
            if next_due_str:
                return datetime.fromisoformat(next_due_str)
        except Exception:
            pass
    now = datetime.now().replace(second=0, microsecond=0)
    next_due = next_run_time(cron_spec, now - timedelta(minutes=1))
    data = {
        "next_due": next_due.isoformat(),
        "last_run": None,
    }
    with state_path.open("w", encoding="utf-8") as handle:
        json.dump(data, handle)
    return next_due


def update_schedule_state(cron_spec: str, state_path: Path) -> tuple[bool, datetime]:
    state_path.parent.mkdir(parents=True, exist_ok=True)
    now = datetime.now().replace(second=0, microsecond=0)
    data: dict
    if state_path.exists():
        try:
            with state_path.open("r", encoding="utf-8") as handle:
                data = json.load(handle)
        except Exception:
            data = {}
    else:
        data = {}

    next_due_str = data.get("next_due")
    if next_due_str:
        try:
            next_due = datetime.fromisoformat(next_due_str)
        except ValueError:
            next_due = next_run_time(cron_spec, now - timedelta(minutes=1))
    else:
        next_due = next_run_time(cron_spec, now - timedelta(minutes=1))

    due = now >= next_due
    if due:
        next_after = next_run_time(cron_spec, next_due)
        data["last_run"] = now.isoformat()
        data["next_due"] = next_after.isoformat()
        with state_path.open("w", encoding="utf-8") as handle:
            json.dump(data, handle)
        return True, next_after

    # Not due yet, ensure next_due is recorded
    data.setdefault("next_due", next_due.isoformat())
    if "last_run" not in data:
        data["last_run"] = None
    with state_path.open("w", encoding="utf-8") as handle:
        json.dump(data, handle)
    return False, next_due


def remove_schedule_state(state_path: Path) -> None:
    try:
        state_path.unlink()
    except FileNotFoundError:
        pass


def remove_schedule_state_by_label(label_seed: str) -> int:
    state_dir = schedule_state_dir()
    count = 0
    for path in state_dir.glob(f"{label_seed}-*.json"):
        try:
            path.unlink()
            count += 1
        except FileNotFoundError:
            continue
    return count


def extract_state_path_from_arguments(arguments: Sequence[str]) -> Path | None:
    for index, value in enumerate(arguments):
        if value == "--schedule-state" and index + 1 < len(arguments):
            return Path(arguments[index + 1]).expanduser()
    return None


def extract_state_path_from_command_line(command_line: str) -> Path | None:
    try:
        tokens = shlex.split(command_line)
    except ValueError:
        return None
    return extract_state_path_from_arguments(tokens)


def prepare_schedule(args: argparse.Namespace) -> tuple[str, str, tzinfo, tzinfo]:
    if not args.schedule:
        raise ValueError("Schedule string is required.")
    schedule_text = args.schedule.strip()
    local_zone = datetime.now().astimezone().tzinfo or timezone.utc
    tz_hint = resolve_schedule_timezone(args.schedule_timezone)
    cron_spec, display_schedule = normalize_schedule(schedule_text, tz_hint, local_zone)
    macro_map = {
        "@yearly": "0 0 1 1 *",
        "@annually": "0 0 1 1 *",
        "@monthly": "0 0 1 * *",
        "@weekly": "0 0 * * 0",
        "@daily": "0 0 * * *",
        "@midnight": "0 0 * * *",
        "@hourly": "0 * * * *",
    }
    cron_spec = macro_map.get(cron_spec, cron_spec)
    return cron_spec, display_schedule, local_zone, tz_hint


def _parse_numeric_field(field: str, name: str, *, allow_wildcard: bool = True) -> int | None:
    if field == "*":
        if allow_wildcard:
            return None
        raise ValueError(f"{name} must be a fixed value for this scheduler backend.")
    if not field.isdigit():
        raise ValueError(
            f"{name} value '{field}' is not supported. Only single integer values are allowed for this backend."
        )
    return int(field)


def cron_to_launchd_interval(cron_spec: str) -> dict:
    minute, hour, day, month, weekday = cron_spec.split()
    minute_value = _parse_numeric_field(minute, "minute", allow_wildcard=False)
    hour_value = _parse_numeric_field(hour, "hour", allow_wildcard=False)
    day_value = _parse_numeric_field(day, "day")
    month_value = _parse_numeric_field(month, "month")
    weekday_value = _parse_numeric_field(weekday, "weekday")

    if day_value is not None and weekday_value is not None:
        raise ValueError(
            "launchd backend does not support combining specific days-of-month and weekdays simultaneously."
        )

    interval: dict[str, int] = {"Minute": minute_value, "Hour": hour_value}
    if day_value is not None:
        interval["Day"] = day_value
    if month_value is not None:
        interval["Month"] = month_value
    if weekday_value is not None:
        interval["Weekday"] = weekday_value
    return interval


def cron_to_systemd_calendar(cron_spec: str) -> str:
    minute, hour, day, month, weekday = cron_spec.split()
    minute_value = _parse_numeric_field(minute, "minute", allow_wildcard=False)
    hour_value = _parse_numeric_field(hour, "hour", allow_wildcard=False)
    day_value = _parse_numeric_field(day, "day")
    month_value = _parse_numeric_field(month, "month")
    weekday_value = _parse_numeric_field(weekday, "weekday")

    if day_value is not None and weekday_value is not None:
        raise ValueError(
            "systemd backend cannot express both a specific day-of-month and weekday at the same time."
        )

    def pad(value: int | None) -> str:
        if value is None:
            return "*"
        return f"{value:02d}"

    date_part = f"*-{pad(month_value)}-{pad(day_value)}"
    time_part = f"{hour_value:02d}:{minute_value:02d}:00"

    if weekday_value is not None:
        weekday_map = {
            0: "Sun",
            1: "Mon",
            2: "Tue",
            3: "Wed",
            4: "Thu",
            5: "Fri",
            6: "Sat",
            7: "Sun",
        }
        if weekday_value not in weekday_map:
            raise ValueError("Weekday must be between 0 (Sunday) and 7 (Sunday).")
        return f"{weekday_map[weekday_value]} *-{pad(month_value)}-* {time_part}"

    return f"*-{pad(month_value)}-{pad(day_value)} {time_part}"


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
    cron_spec, display_schedule, local_zone, _ = prepare_schedule(args)

    label_seed = args.schedule_label or args.csv.stem or "mailmerge"
    label_seed = sanitize_label_seed(label_seed)

    base_command = build_program_arguments(args)
    fingerprint_source = "\n".join(base_command + [cron_spec])
    fingerprint = hashlib.sha1(fingerprint_source.encode("utf-8")).hexdigest()[:8]
    state_path = build_state_path(label_seed, fingerprint)
    next_due = initialize_schedule_state(state_path, cron_spec, overwrite=args.schedule_overwrite)

    command = build_cron_command(args, schedule_spec=cron_spec, state_path=state_path)
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
    tz_note = f" interpreted in {args.schedule_timezone}" if args.schedule_timezone else ""
    logging.info(
        "Installed cron entry '%s' (%s%s) using cron spec '%s' (system timezone %s).",
        label_seed,
        display_schedule,
        tz_note,
        cron_spec,
        system_zone_label,
    )
    logging.info("Next run scheduled for %s.", next_due.isoformat())
    if not args.password and not os.getenv("GMAIL_APP_PASSWORD"):
        logging.warning(
            "Cron job does not include a password. Ensure the job environment defines GMAIL_APP_PASSWORD."
        )
    if not args.sender and not os.getenv("GMAIL_ADDRESS"):
        logging.warning(
            "Cron job does not include a sender address. Ensure the job environment defines GMAIL_ADDRESS."
        )


def install_launchd_job(args: argparse.Namespace) -> None:
    require_backend_supported("launchd")
    cron_spec, display_schedule, local_zone, _ = prepare_schedule(args)
    interval = cron_to_launchd_interval(cron_spec)

    label_seed_raw = args.schedule_label or args.csv.stem or "mailmerge"
    label_seed = sanitize_label_seed(label_seed_raw)
    base_command = build_program_arguments(args)
    fingerprint_source = "\n".join(base_command + [cron_spec])
    fingerprint = hashlib.sha1(fingerprint_source.encode("utf-8")).hexdigest()[:8]
    state_path = build_state_path(label_seed, fingerprint)
    next_due = initialize_schedule_state(state_path, cron_spec, overwrite=args.schedule_overwrite)

    label = f"{LAUNCHD_LABEL_PREFIX}{label_seed}"

    launch_agents_dir = Path.home() / "Library" / "LaunchAgents"
    launch_agents_dir.mkdir(parents=True, exist_ok=True)
    plist_path = launch_agents_dir / f"{label}.plist"

    program_args = build_program_arguments(args, schedule_spec=cron_spec, state_path=state_path)
    log_dir = Path.home() / "Library" / "Logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"{label}.log"

    if plist_path.exists():
        if not args.schedule_overwrite:
            raise RuntimeError(
                f"LaunchAgent '{label}' already exists. Use --schedule-overwrite to replace it."
            )
        subprocess.run(
            ["launchctl", "bootout", f"gui/{os.getuid()}", str(plist_path)],
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

    payload = {
        "Label": label,
        "ProgramArguments": program_args,
        "RunAtLoad": True,
        "StartCalendarInterval": interval,
        "StandardOutPath": str(log_path),
        "StandardErrorPath": str(log_path),
        "WorkingDirectory": str(args.csv.parent),
    }

    with plist_path.open("wb") as handle:
        plistlib.dump(payload, handle)

    load_result = subprocess.run(
        ["launchctl", "bootstrap", f"gui/{os.getuid()}", str(plist_path)],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if load_result.returncode != 0:
        raise RuntimeError(
            f"Failed to load launchd job: {load_result.stderr.strip() or load_result.stdout.strip()}"
        )

    system_zone_label = describe_timezone(local_zone)
    tz_note = f" interpreted in {args.schedule_timezone}" if args.schedule_timezone else ""
    logging.info(
        "Installed launchd job '%s' (%s%s) with StartCalendarInterval %s (system timezone %s).",
        label,
        display_schedule,
        tz_note,
        interval,
        system_zone_label,
    )
    logging.info("Next run scheduled for %s.", next_due.isoformat())

    if not args.password and not os.getenv("GMAIL_APP_PASSWORD"):
        logging.warning(
            "Launchd job does not include a password. Ensure the environment provides GMAIL_APP_PASSWORD."
        )
    if not args.sender and not os.getenv("GMAIL_ADDRESS"):
        logging.warning(
            "Launchd job does not include a sender address. Ensure the environment provides GMAIL_ADDRESS."
        )


def install_systemd_job(args: argparse.Namespace) -> None:
    require_backend_supported("systemd")
    if shutil.which("systemctl") is None:
        raise RuntimeError("systemctl not found. The systemd backend requires systemctl to manage timers.")

    cron_spec, display_schedule, local_zone, _ = prepare_schedule(args)
    calendar_spec = cron_to_systemd_calendar(cron_spec)

    label_seed_raw = args.schedule_label or args.csv.stem or "mailmerge"
    label_seed = sanitize_label_seed(label_seed_raw)
    base_command = build_program_arguments(args)
    fingerprint_source = "\n".join(base_command + [cron_spec])
    fingerprint = hashlib.sha1(fingerprint_source.encode("utf-8")).hexdigest()[:8]
    state_path = build_state_path(label_seed, fingerprint)
    next_due = initialize_schedule_state(state_path, cron_spec, overwrite=args.schedule_overwrite)

    unit_name = f"{SYSTEMD_UNIT_PREFIX}{label_seed}"

    systemd_dir = Path.home() / ".config" / "systemd" / "user"
    systemd_dir.mkdir(parents=True, exist_ok=True)

    service_path = systemd_dir / f"{unit_name}.service"
    timer_path = systemd_dir / f"{unit_name}.timer"

    program_args = build_program_arguments(args, schedule_spec=cron_spec, state_path=state_path)
    exec_start = shlex.join(program_args)

    service_unit = "\n".join(
        [
            "[Unit]",
            f"Description=Mailmerge job {label_seed}",
            "",
            "[Service]",
            "Type=oneshot",
            f"WorkingDirectory={args.csv.parent}",
            f"ExecStart={exec_start}",
            "",
            "[Install]",
            "WantedBy=default.target",
            "",
        ]
    )

    timer_unit = "\n".join(
        [
            "[Unit]",
            f"Description=Mailmerge timer {label_seed}",
            "",
            "[Timer]",
            f"OnCalendar={calendar_spec}",
            "Persistent=true",
            f"Unit={unit_name}.service",
            "",
            "[Install]",
            "WantedBy=timers.target",
            "",
        ]
    )

    if (service_path.exists() or timer_path.exists()) and not args.schedule_overwrite:
        raise RuntimeError(
            f"Systemd units for '{unit_name}' already exist. Use --schedule-overwrite to replace them."
        )

    if timer_path.exists():
        subprocess.run(
            ["systemctl", "--user", "disable", "--now", f"{unit_name}.timer"],
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

    with service_path.open("w", encoding="utf-8") as handle:
        handle.write(service_unit)
    with timer_path.open("w", encoding="utf-8") as handle:
        handle.write(timer_unit)

    reload_result = subprocess.run(
        ["systemctl", "--user", "daemon-reload"],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if reload_result.returncode != 0:
        raise RuntimeError(
            f"Failed to reload systemd user units: {reload_result.stderr.strip() or reload_result.stdout.strip()}"
        )

    enable_result = subprocess.run(
        ["systemctl", "--user", "enable", "--now", f"{unit_name}.timer"],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if enable_result.returncode != 0:
        raise RuntimeError(
            f"Failed to enable systemd timer: {enable_result.stderr.strip() or enable_result.stdout.strip()}"
        )

    system_zone_label = describe_timezone(local_zone)
    tz_note = f" interpreted in {args.schedule_timezone}" if args.schedule_timezone else ""
    logging.info(
        "Installed systemd timer '%s' (%s%s) with OnCalendar '%s' (system timezone %s, Persistent=true).",
        unit_name,
        display_schedule,
        tz_note,
        calendar_spec,
        system_zone_label,
    )
    logging.info("Next run scheduled for %s.", next_due.isoformat())

    if not args.password and not os.getenv("GMAIL_APP_PASSWORD"):
        logging.warning(
            "Systemd job does not include a password. Ensure the job environment defines GMAIL_APP_PASSWORD."
        )
    if not args.sender and not os.getenv("GMAIL_ADDRESS"):
        logging.warning(
            "Systemd job does not include a sender address. Ensure the job environment defines GMAIL_ADDRESS."
        )


def remove_mailmerge_cron_jobs() -> int:
    lines = read_crontab_lines()
    if not lines:
        return 0

    new_lines: List[str] = []
    removed_jobs = 0
    index = 0
    while index < len(lines):
        line = lines[index]
        if line.startswith(CRON_COMMENT_PREFIX):
            match = re.match(rf"{re.escape(CRON_COMMENT_PREFIX)}\s+([^\s(]+)(?:\s+\(([0-9a-fA-F]+)\))?", line)
            label_seed = None
            fingerprint = None
            if match:
                label_seed = match.group(1)
                fingerprint = match.group(2)

            command_line = lines[index + 1] if index + 1 < len(lines) else ""
            state_path = extract_state_path_from_command_line(command_line)
            if state_path is not None:
                remove_schedule_state(state_path)
            elif label_seed is not None:
                remove_schedule_state_by_label(label_seed)

            index += 2
            removed_jobs += 1
            if new_lines and not new_lines[-1].strip():
                new_lines.pop()
            continue

        new_lines.append(line)
        index += 1

    while new_lines and not new_lines[-1].strip():
        new_lines.pop()

    if removed_jobs > 0:
        write_crontab_lines(new_lines)

    return removed_jobs


def remove_launchd_jobs() -> int:
    require_backend_supported("launchd")
    launch_agents_dir = Path.home() / "Library" / "LaunchAgents"
    if not launch_agents_dir.exists():
        return 0

    removed = 0
    uid = os.getuid()
    for plist_path in launch_agents_dir.glob(f"{LAUNCHD_LABEL_PREFIX}*.plist"):
        state_path = None
        if plist_path.stem.startswith(LAUNCHD_LABEL_PREFIX):
            label_seed = plist_path.stem[len(LAUNCHD_LABEL_PREFIX) :]
        else:
            label_seed = plist_path.stem
        try:
            with plist_path.open("rb") as handle:
                payload = plistlib.load(handle)
            program_args = payload.get("ProgramArguments", []) or []
            state_path_candidate = extract_state_path_from_arguments(program_args)
            if state_path_candidate is not None:
                state_path = state_path_candidate
        except Exception:
            state_path = None
        subprocess.run(
            ["launchctl", "bootout", f"gui/{uid}", str(plist_path)],
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
        try:
            plist_path.unlink()
            removed += 1
        except FileNotFoundError:
            continue
        if state_path is not None:
            remove_schedule_state(state_path)
        else:
            remove_schedule_state_by_label(label_seed)
    return removed


def remove_systemd_jobs() -> int:
    require_backend_supported("systemd")
    systemd_dir = Path.home() / ".config" / "systemd" / "user"
    if not systemd_dir.exists():
        return 0

    timers = list(systemd_dir.glob(f"{SYSTEMD_UNIT_PREFIX}*.timer"))
    if not timers:
        return 0

    has_systemctl = shutil.which("systemctl") is not None
    removed = 0

    for timer_path in timers:
        unit_name = timer_path.stem
        if unit_name.startswith(SYSTEMD_UNIT_PREFIX):
            label_seed = unit_name[len(SYSTEMD_UNIT_PREFIX) :]
        else:
            label_seed = unit_name
        if has_systemctl:
            subprocess.run(
                ["systemctl", "--user", "disable", "--now", f"{unit_name}.timer"],
                check=False,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
        service_path = systemd_dir / f"{unit_name}.service"
        state_path = None
        if service_path.exists():
            try:
                content = service_path.read_text(encoding="utf-8")
                for line in content.splitlines():
                    if line.startswith("ExecStart="):
                        command_line = line.partition("=")[2].strip()
                        state_candidate = extract_state_path_from_command_line(command_line)
                        if state_candidate is not None:
                            state_path = state_candidate
                        break
            except Exception:
                state_path = None
        try:
            timer_path.unlink()
            removed += 1
        except FileNotFoundError:
            continue
        if service_path.exists():
            service_path.unlink()
        if state_path is not None:
            remove_schedule_state(state_path)
        else:
            remove_schedule_state_by_label(label_seed)

    if removed and has_systemctl:
        subprocess.run(
            ["systemctl", "--user", "daemon-reload"],
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
    return removed


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
    cc_templates: Sequence[Template],
    bcc_templates: Sequence[Template],
    cc_column: str | None,
    bcc_column: str | None,
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

            try:
                cc_addresses: List[str] = []
                for template in cc_templates:
                    rendered_cc = format_message(
                        template,
                        context,
                        template_label="cc",
                    )
                    cc_addresses.extend(parse_addresses(rendered_cc))
            except ValueError as exc:
                logging.error("Skipping row %s: %s", index, exc)
                continue

            if cc_column:
                cc_raw = row.get(cc_column)
                if isinstance(cc_raw, str):
                    cc_addresses.extend(parse_addresses(cc_raw))
                elif cc_raw:
                    cc_addresses.extend(parse_addresses(str(cc_raw)))

            try:
                bcc_addresses: List[str] = []
                for template in bcc_templates:
                    rendered_bcc = format_message(
                        template,
                        context,
                        template_label="bcc",
                    )
                    bcc_addresses.extend(parse_addresses(rendered_bcc))
            except ValueError as exc:
                logging.error("Skipping row %s: %s", index, exc)
                continue

            if bcc_column:
                bcc_raw = row.get(bcc_column)
                if isinstance(bcc_raw, str):
                    bcc_addresses.extend(parse_addresses(bcc_raw))
                elif bcc_raw:
                    bcc_addresses.extend(parse_addresses(str(bcc_raw)))

            if cc_addresses:
                cc_addresses = list(dict.fromkeys(cc_addresses))
            if bcc_addresses:
                bcc_addresses = list(dict.fromkeys(bcc_addresses))

            message = build_message(
                sender,
                recipients,
                subject,
                body,
                subtype=subtype,
                reply_to=reply_to,
                cc=cc_addresses,
                bcc=bcc_addresses,
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

    remove_requested = bool(args.schedule_remove_all)
    backend = determine_backend(args.schedule_backend)

    if remove_requested and args.schedule:
        logging.error("Cannot combine --schedule with --schedule-remove-all.")
        return 1

    if remove_requested and (args.schedule_label or args.schedule_overwrite):
        logging.error("--schedule-remove-all cannot be combined with other scheduling flags.")
        return 1

    if remove_requested:
        try:
            if backend == "cron":
                removed = remove_mailmerge_cron_jobs()
            elif backend == "launchd":
                removed = remove_launchd_jobs()
            else:
                removed = remove_systemd_jobs()
        except Exception as exc:
            logging.error("%s", exc)
            return 1
        if removed:
            plural = "entry" if removed == 1 else "entries"
            logging.info(
                "Removed %s scheduled mailmerge %s %s.",
                removed,
                backend,
                plural,
            )
        else:
            logging.info("No mailmerge %s entries found to remove.", backend)
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
            if backend == "cron":
                install_cron_job(args)
            elif backend == "launchd":
                install_launchd_job(args)
            else:
                install_systemd_job(args)
        except Exception as exc:
            logging.error("%s", exc)
            return 1
        return 0

    if args.schedule_state and args.schedule_spec:
        try:
            due, next_due = update_schedule_state(args.schedule_spec, args.schedule_state)
        except Exception as exc:
            logging.error("Failed to evaluate schedule state: %s", exc)
            return 1
        if not due:
            logging.info("No messages due. Next run scheduled for %s.", next_due.isoformat())
            return 0
        logging.info("Scheduled time reached; proceeding with mail merge. Next run scheduled for %s.", next_due.isoformat())

    try:
        rows = load_recipients(args.csv)
        subject_template = Template(args.subject)
        body_template = read_template(args.body)
        cc_templates = [Template(value) for value in args.cc]
        bcc_templates = [Template(value) for value in args.bcc]
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
        cc_templates=cc_templates,
        bcc_templates=bcc_templates,
        cc_column=args.cc_column,
        bcc_column=args.bcc_column,
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
