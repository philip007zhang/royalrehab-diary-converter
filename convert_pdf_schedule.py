#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import sys
import uuid
import webbrowser
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Iterable


@dataclass(frozen=True)
class ScheduleEvent:
    start: datetime
    end: datetime
    title: str
    location: str = ""


_LINE_PATTERNS = (
    re.compile(
        r"^(?P<date>\d{4}-\d{2}-\d{2})\s+(?P<start>\d{1,2}:\d{2})\s*-\s*(?P<end>\d{1,2}:\d{2})\s+(?P<title>.+?)(?:\s+@\s+(?P<location>.+))?$"
    ),
    re.compile(
        r"^(?P<date>\d{2}/\d{2}/\d{4})\s+(?P<start>\d{1,2}:\d{2})\s*-\s*(?P<end>\d{1,2}:\d{2})\s+(?P<title>.+?)(?:\s+@\s+(?P<location>.+))?$"
    ),
)


def _escape_ics_value(value: str) -> str:
    return (
        value.replace("\\", "\\\\")
        .replace(";", r"\;")
        .replace(",", r"\,")
        .replace("\n", r"\n")
    )


def parse_schedule_text(schedule_text: str) -> list[ScheduleEvent]:
    events: list[ScheduleEvent] = []
    for raw_line in schedule_text.splitlines():
        line = " ".join(raw_line.split()).strip()
        if not line:
            continue
        for pattern in _LINE_PATTERNS:
            match = pattern.match(line)
            if not match:
                continue
            date_raw = match.group("date")
            date_format = "%Y-%m-%d" if "-" in date_raw else "%d/%m/%Y"
            start = datetime.strptime(
                f"{date_raw} {match.group('start')}", f"{date_format} %H:%M"
            )
            end = datetime.strptime(
                f"{date_raw} {match.group('end')}", f"{date_format} %H:%M"
            )
            events.append(
                ScheduleEvent(
                    start=start,
                    end=end,
                    title=match.group("title").strip(),
                    location=(match.group("location") or "").strip(),
                )
            )
            break
    events.sort(key=lambda item: item.start)
    return events


def build_ics(events: Iterable[ScheduleEvent], timezone: str = "UTC") -> str:
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Royal Rehab//Diary Converter//EN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
    ]
    stamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%SZ")
    for event in events:
        lines.extend(
            [
                "BEGIN:VEVENT",
                f"UID:{uuid.uuid4()}",
                f"DTSTAMP:{stamp}",
                f"DTSTART;TZID={timezone}:{event.start.strftime('%Y%m%dT%H%M%S')}",
                f"DTEND;TZID={timezone}:{event.end.strftime('%Y%m%dT%H%M%S')}",
                f"SUMMARY:{_escape_ics_value(event.title)}",
                f"LOCATION:{_escape_ics_value(event.location)}",
                "END:VEVENT",
            ]
        )
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines) + "\r\n"


def extract_text_from_pdf(pdf_path: Path) -> str:
    try:
        from pypdf import PdfReader
    except ModuleNotFoundError as exc:  # pragma: no cover - environment-dependent
        raise RuntimeError(
            "Missing dependency 'pypdf'. Install it with: pip install pypdf"
        ) from exc

    reader = PdfReader(str(pdf_path))
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def convert_pdf_to_ics(input_pdf: Path, output_ics: Path, timezone: str) -> int:
    schedule_text = extract_text_from_pdf(input_pdf)
    events = parse_schedule_text(schedule_text)
    if not events:
        print(
            "No schedule rows were detected. Expected lines like "
            "'2026-05-13 09:00-10:00 Physiotherapy @ Royal Rehab'.",
            file=sys.stderr,
        )
        return 1
    output_ics.write_text(build_ics(events, timezone=timezone), encoding="utf-8")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert a PDF schedule into an .ics calendar file."
    )
    parser.add_argument("input_pdf", type=Path, help="Path to the source PDF schedule")
    parser.add_argument("output_ics", type=Path, help="Path for generated .ics file")
    parser.add_argument(
        "--timezone",
        default="UTC",
        help="IANA timezone name for generated calendar events (default: UTC)",
    )
    parser.add_argument(
        "--open-calendar",
        action="store_true",
        help="Open the generated .ics file with the system calendar app",
    )
    args = parser.parse_args()

    exit_code = convert_pdf_to_ics(args.input_pdf, args.output_ics, args.timezone)
    if exit_code != 0:
        return exit_code

    print(f"Generated calendar file: {args.output_ics}")
    if args.open_calendar:
        webbrowser.open(args.output_ics.resolve().as_uri())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
