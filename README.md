# Royal Rehab Schedule Converter

Small Node.js app for turning a Royal Rehab patient diary image or PDF into portable calendar files.

## What it does

- Uploads a diary screenshot, export image, or PDF.
- Runs OCR in the browser with Tesseract.js.
- Extracts appointment rows using the report's column layout.
- Lets you review and edit sessions before export.
- Downloads `.ics`, CSV, or JSON output.

## Run it

```powershell
node server.js
```

Then open [http://127.0.0.1:3000](http://127.0.0.1:3000).

To expose the app on a network interface, set `HOST` and `PORT` when starting it, for example `HOST=0.0.0.0 PORT=7000`.

## Optional Google Calendar direct add

To enable one-step Google Calendar insertion instead of downloading and importing an `.ics` file, start the app with:

```powershell
$env:GOOGLE_CLIENT_ID="your-google-oauth-client-id"
node server.js
```

When `GOOGLE_CLIENT_ID` is present, the app shows `Connect Google` and can add the merged appointments directly to the signed-in user's primary Google Calendar after one confirmation.

If you already have a Google API key configured, the app still accepts it, but it is no longer required for direct add.

For non-local deployments, Google OAuth also requires the app's exact origin to be added to the OAuth client's `Authorized JavaScript origins`, for example `http://47.94.17.128:7000`.

## Audit log

The app now records conversion activity into `audit.log.jsonl` and exposes an admin-only audit view at [http://127.0.0.1:3000/auditlog](http://127.0.0.1:3000/auditlog).

To set your own admin login:

```powershell
$env:ADMIN_USERNAME="your-admin-user"
$env:ADMIN_PASSWORD="your-strong-password"
node server.js
```

If those values are not set, the server starts with default audit log credentials and prints them to the console on startup.

## Notes

- This version is tuned for the Royal Rehab diary layout shown in `Sample-schedule.png`.
- OCR is loaded from a CDN on first use, so the browser needs internet access.
- Supported inputs: PNG, JPG, WEBP, BMP, and PDF.
- PDF uploads are rendered in the browser and OCR runs page by page, with the first page shown in the preview area.
