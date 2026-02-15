# Firepulse PDF Extraction Service

Flask-based microservice for extracting text from PDF files.

## API Endpoints

### POST /extract-pdf

Extract text from a PDF file.

**Headers:**
- `X-API-Key`: Your API key
- `Content-Type`: application/json

**Body:**
```json
{
  "file_base64": "base64_encoded_pdf_content"
}
```

**Response:**
```json
{
  "success": true,
  "text": "Full extracted text...",
  "total_pages": 2,
  "total_chars": 6543,
  "pages": [...]
}
```

## Environment Variables

- `API_KEY`: API key for authentication (default: fp_pdf_service_2026_key)
- `PORT`: Port to run the service (default: 5000)

## Local Testing
```bash
pip install -r requirements.txt
python app.py
```

## Deploy to Render

1. Push to GitHub
2. Connect Render to your repo
3. Set environment variable: `API_KEY`
4. Deploy