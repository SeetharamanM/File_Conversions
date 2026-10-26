"""Flask web app for converting Excel, CSV, PDF, Word, Google Docs and Google Sheets files."""

from __future__ import annotations

import io
import os

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.utils import secure_filename

import converters
from converters import ConversionError

MAX_UPLOAD_MB = int(os.environ.get("MAX_UPLOAD_MB", "50"))

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024


@app.get("/")
def index():
    return render_template("index.html", max_mb=MAX_UPLOAD_MB)


@app.get("/api/formats")
def formats():
    return jsonify(converters.format_catalogue())


@app.post("/api/convert")
def convert():
    target = (request.form.get("target") or "").lower().strip()
    url = (request.form.get("url") or "").strip()
    upload = request.files.get("file")

    try:
        if not target:
            raise ConversionError("Choose a format to convert to.")
        if url:
            src, _ = converters.parse_google_url(url)
            data = url.encode("utf-8")
            stem = "google-doc" if src == "gdoc" else "google-sheet"
        elif upload and upload.filename:
            src = converters.detect_format(upload.filename)
            data = upload.read()
            if not data:
                raise ConversionError("The uploaded file is empty.")
            stem = os.path.splitext(secure_filename(upload.filename))[0] or "converted"
        else:
            raise ConversionError("Upload a file or paste a Google Docs / Sheets link.")

        result = converters.convert(data, src, target, stem=stem)
    except ConversionError as exc:
        return jsonify(error=str(exc)), 400

    return send_file(
        io.BytesIO(result.data),
        mimetype=result.mimetype,
        as_attachment=True,
        download_name=result.filename,
    )


@app.errorhandler(RequestEntityTooLarge)
def too_large(_exc):
    return jsonify(error=f"File is too large (limit {MAX_UPLOAD_MB} MB)."), 413


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")), debug=False)
