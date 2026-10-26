FROM python:3.12-slim

# LibreOffice enables .doc input/output and high-fidelity Word/Excel -> PDF.
RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-writer libreoffice-calc fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

ENV PORT=8080 PYTHONUNBUFFERED=1
EXPOSE 8080
# LibreOffice conversions can take a while, so allow long requests.
CMD exec gunicorn --bind 0.0.0.0:$PORT --workers 2 --threads 4 --timeout 180 app:app
