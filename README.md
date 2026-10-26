# File Converter

A small web app that converts files between formats. Upload a file (or paste a
Google Docs / Sheets link), pick a target format, and the converted file downloads.

| From | To |
|---|---|
| Excel (.xlsx, .xls), CSV, Google Sheets | xlsx, csv, pdf, docx, json, html, txt, md |
| Word (.docx, .doc), Google Docs | pdf, docx, doc, txt, html, md, xlsx, csv |
| PDF | docx, txt, html, md, xlsx, csv, json |

Notes:
- **Google Docs / Sheets**: paste a link, or upload a `.gdoc` / `.gsheet` shortcut
  from Google Drive for desktop. The file must be shared as *Anyone with the link can view*;
  the app downloads it through Google's export endpoint, then converts it.
- **LibreOffice** (optional, `soffice` on PATH) is used for `.doc` input/output and for
  high-fidelity Word/Excel → PDF. Without it those fall back to a simpler layout and
  `.doc` is unavailable.
- **PDF → Excel/CSV** uses detected tables; if a PDF has none, each text line becomes a row.
  Scanned (image-only) PDFs need OCR and aren't supported.
- A workbook with several sheets converted to CSV downloads as a ZIP of one CSV per sheet.

## Run

```bash
pip install -r requirements.txt
python app.py            # http://localhost:5000  (PORT, MAX_UPLOAD_MB env vars)
```

## Deploy

The app ships with a `Dockerfile` (Python + LibreOffice + gunicorn) and a `Procfile`.
Use Docker-based hosting if you want `.doc` support and high-fidelity PDFs.

**Docker (any server / VPS)**
```bash
docker build -t file-converter .
docker run -d -p 80:8080 --restart unless-stopped file-converter
```

**Google Cloud Run**
```bash
gcloud run deploy file-converter --source . --region us-central1 \
  --allow-unauthenticated --memory 2Gi --timeout 300
```

**Render / Railway / Fly.io**: connect the GitHub repo and choose the Docker
runtime; it builds from the `Dockerfile` and listens on `$PORT`.
(A plain Python/Procfile runtime works too, but without LibreOffice.)

Settings: `PORT` (default 8080 in Docker), `MAX_UPLOAD_MB` (default 50).
Give it at least 1–2 GB RAM: LibreOffice and large spreadsheets use memory.

## Test

```bash
python -m pytest -q
```

## Layout

- `converters.py` – conversion engine (readers → intermediate table/document model → writers)
- `app.py` – Flask routes: `/`, `GET /api/formats`, `POST /api/convert` (`file` or `url`, plus `target`)
- `templates/`, `static/` – the web UI
