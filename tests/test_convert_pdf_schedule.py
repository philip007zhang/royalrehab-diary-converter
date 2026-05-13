import unittest
from datetime import datetime

from convert_pdf_schedule import ScheduleEvent, build_ics, parse_schedule_text


class ParseScheduleTextTests(unittest.TestCase):
    def test_parses_multiple_supported_date_formats(self) -> None:
        content = """
        2026-05-13 09:00-10:00 Physiotherapy @ Royal Rehab
        14/05/2026 11:30-12:15 Hydro Therapy @ Pool
        """

        events = parse_schedule_text(content)

        self.assertEqual(len(events), 2)
        self.assertEqual(events[0].title, "Physiotherapy")
        self.assertEqual(events[0].location, "Royal Rehab")
        self.assertEqual(events[1].start, datetime(2026, 5, 14, 11, 30))

    def test_ignores_non_matching_lines(self) -> None:
        content = "Not a schedule row\nAnother invalid line"

        events = parse_schedule_text(content)

        self.assertEqual(events, [])


class BuildIcsTests(unittest.TestCase):
    def test_generates_basic_ics_document(self) -> None:
        events = [
            ScheduleEvent(
                start=datetime(2026, 5, 13, 9, 0),
                end=datetime(2026, 5, 13, 10, 0),
                title="Physiotherapy",
                location="Royal Rehab",
            )
        ]

        ics = build_ics(events, timezone="Australia/Sydney")

        self.assertIn("BEGIN:VCALENDAR", ics)
        self.assertIn("BEGIN:VEVENT", ics)
        self.assertIn("DTSTART;TZID=Australia/Sydney:20260513T090000", ics)
        self.assertIn("SUMMARY:Physiotherapy", ics)
        self.assertIn("LOCATION:Royal Rehab", ics)


if __name__ == "__main__":
    unittest.main()
