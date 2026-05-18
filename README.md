# Royal Rehab Schedule Converter

Small Node.js app for turning a Royal Rehab patient diary image or PDF into portable calendar files.

## What it does

- Uploads a diary screenshot, export image, or PDF.
- Runs OCR in the browser with Tesseract.js.
- Falls back to server-side OCR when browser JavaScript or browser OCR is unavailable.
- Extracts appointment rows using the report's column layout.
- Lets you review and edit sessions before export.
- Downloads `.ics`, CSV, or JSON output.

## Run it

```powershell
node server.js
```

Then open [http://127.0.0.1:3000](http://127.0.0.1:3000).

To expose the app on a network interface, set `HOST` and `PORT` when starting it, for example `HOST=0.0.0.0 PORT=7000`.

## Server-side OCR fallback

If browser JavaScript is disabled, the root page shows a plain upload form that posts the file to `/fallback/extract`. The same server OCR path is also used automatically by the app if the browser OCR libraries fail to load.

When the browser app falls back to server OCR, it now tracks a background OCR job and shows live phases such as upload received, PDF rendering, OCR scanning, and finalizing extraction.

Server OCR uses native commands:

- `tesseract` for image OCR
- `pdftoppm` for PDF page rendering before OCR

On Ubuntu or Debian, install them with:

```bash
sudo apt-get update
sudo apt-get install -y tesseract-ocr poppler-utils
```

You can verify that the backend fallback is available by checking:

```text
/api/server-ocr-status
```

The browser-tracked job endpoints are:

```text
POST /api/server-ocr-jobs
GET  /api/server-ocr-jobs/:id
```

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
- Server OCR does not depend on browser JavaScript, but it does require the native commands above on the Node host.
- Supported inputs: PNG, JPG, WEBP, BMP, and PDF.
- PDF uploads are rendered in the browser and OCR runs page by page, with the first page shown in the preview area.
