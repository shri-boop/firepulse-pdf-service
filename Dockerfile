FROM python:3.11-slim

# System deps required by the extractor:
#   tesseract-ocr  -> OCR fallback for scanned PDFs (pytesseract)
#   poppler-utils  -> pdf2image (rasterising PDF pages for OCR)
# (mirrors the buildpack 'aptfile')
RUN apt-get update && apt-get install -y \
    tesseract-ocr \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Listens on 5000 inside the container. The API key is read from the API_KEY env
# var (defaults in app.py to the shared key the n8n workflow already sends).
EXPOSE 5000

CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--timeout", "300", "--workers", "2", "app:app"]
