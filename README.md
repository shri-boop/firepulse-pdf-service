# Firepulse Document Extraction Service

Comprehensive microservice for extracting text from PDFs and Excel files.

## Features

### PDF Extraction
- Text extraction with layout preservation
- Table structure extraction
- OCR for scanned PDFs
- Multi-column layout handling

### Excel Extraction (.xls, .xlsx)
- Multi-sheet extraction
- Table structure preservation
- Formula and value extraction
- Merged cell handling
- Date and number formatting

## API Endpoints

### POST /extract-pdf

Extract text from PDF files.

**Request:**
```json
{
  "file_base64": "base64_encoded_pdf_content"
}
```

**Response:**
```json
{
  "success": true,
  "text": "Extracted text...",
  "total_pages": 2,
  "total_chars": 6543,
  "extraction_method": "pdfplumber"
}
```

### POST /extract-excel

Extract text from Excel files (.xls, .xlsx).

**Request:**
```json
{
  "file_base64": "base64_encoded_excel_content",
  "file_extension": ".xlsx"
}
```

**Response:**
```json
{
  "success": true,
  "text": "Extracted text...",
  "total_sheets": 3,
  "total_chars": 4521,
  "sheets": [...]
}
```

## Environment Variables

- `API_KEY`: API key for authentication
- `PORT`: Port to run the service (default: 5000)

## Deployment

Designed for Render.com deployment with automatic detection.