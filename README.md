# Royal Rehab Schedule Converter

Small Node.js app for turning a Royal Rehab patient diary image into portable calendar files.

## What it does

- Uploads a diary screenshot or export image.
- Runs OCR in the browser with Tesseract.js.
- Extracts appointment rows using the report's column layout.
- Lets you review and edit sessions before export.
- Downloads `.ics`, CSV, or JSON output.

## Run it

```powershell
node server.js
```

Then open [http://127.0.0.1:3000](http://127.0.0.1:3000).

## Optional Google Calendar direct add

To enable one-step Google Calendar insertion instead of downloading and importing an `.ics` file, start the app with:

```powershell
$env:GOOGLE_CLIENT_ID="your-google-oauth-client-id"
$env:GOOGLE_API_KEY="your-google-api-key"
node server.js
```

When those values are present, the app shows `Connect Google` and can add the merged appointments directly to the signed-in user's primary Google Calendar after one confirmation.

## Notes

- This version is tuned for the Royal Rehab diary layout shown in `Sample-schedule.png`.
- OCR is loaded from a CDN on first use, so the browser needs internet access.
- Image inputs are supported in this first version: PNG, JPG, WEBP, and BMP.
