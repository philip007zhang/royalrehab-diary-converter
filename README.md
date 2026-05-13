# royalrehab-diary-converter

Convert Royal Rehab schedule PDFs into calendar events.

## Convert PDF schedule to `.ics`

1. Install dependency:

   ```bash
   pip install pypdf
   ```

2. Run the converter:

   ```bash
   python convert_pdf_schedule.py /path/to/schedule.pdf /path/to/schedule.ics --timezone Australia/Sydney
   ```

3. Optional: open the generated `.ics` directly in your default calendar app:

   ```bash
   python convert_pdf_schedule.py /path/to/schedule.pdf /path/to/schedule.ics --open-calendar
   ```

## Expected PDF text format

The converter extracts text from the PDF and looks for rows like:

- `2026-05-13 09:00-10:00 Physiotherapy @ Royal Rehab`
- `13/05/2026 09:00-10:00 Physiotherapy @ Royal Rehab`
