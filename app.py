from flask import Flask, request, jsonify
from flask_cors import CORS
import pdfplumber
import base64
import io
import os
from functools import wraps
from PIL import Image
import pytesseract
from pdf2image import convert_from_bytes

app = Flask(__name__)
CORS(app)

# =============================
# SECURITY - API KEY
# =============================

API_KEY = os.environ.get('API_KEY', 'fp_pdf_service_2026_key')

def require_api_key(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        api_key = request.headers.get('X-API-Key')
        
        if not api_key or api_key != API_KEY:
            return jsonify({
                'error': 'Unauthorized',
                'message': 'Invalid or missing API key'
            }), 401
        
        return f(*args, **kwargs)
    
    return decorated_function

# =============================
# HEALTH CHECK
# =============================

@app.route('/', methods=['GET'])
def health_check():
    return jsonify({
        'status': 'active',
        'service': 'Firepulse PDF Extraction Service (Enhanced)',
        'version': '2.0.0',
        'capabilities': [
            'Text extraction with layout preservation',
            'Table structure extraction',
            'OCR for scanned PDFs',
            'Multi-column layout handling'
        ]
    })

# =============================
# HELPER: OCR FOR SCANNED PDFs
# =============================

def extract_with_ocr(pdf_bytes):
    """Extract text from scanned PDF using OCR"""
    try:
        # Convert PDF to images
        images = convert_from_bytes(pdf_bytes)
        
        full_text = ''
        page_texts = []
        
        for page_num, image in enumerate(images):
            # Perform OCR on each page
            page_text = pytesseract.image_to_string(image)
            
            page_texts.append({
                'page_number': page_num + 1,
                'text': page_text,
                'char_count': len(page_text),
                'method': 'OCR'
            })
            
            full_text += page_text + '\n'
        
        return full_text, page_texts
    
    except Exception as e:
        raise Exception(f'OCR extraction failed: {str(e)}')

# =============================
# HELPER: EXTRACT TABLES
# =============================

def extract_tables_from_page(page):
    """Extract tables and convert to readable text"""
    tables = page.extract_tables()
    
    if not tables:
        return ""
    
    table_text = ""
    
    for table_num, table in enumerate(tables):
        table_text += f"\n[TABLE {table_num + 1}]\n"
        
        for row in table:
            # Join cells with | separator
            row_text = " | ".join([str(cell) if cell else "" for cell in row])
            table_text += row_text + "\n"
        
        table_text += "[END TABLE]\n"
    
    return table_text

# =============================
# PDF EXTRACTION ENDPOINT
# =============================

@app.route('/extract-pdf', methods=['POST'])
@require_api_key
def extract_pdf():
    try:
        # Get JSON data
        data = request.json
        
        if not data or 'file_base64' not in data:
            return jsonify({
                'error': 'Bad Request',
                'message': 'Missing file_base64 in request body'
            }), 400
        
        # Decode base64
        pdf_base64 = data['file_base64']
        
        try:
            pdf_bytes = base64.b64decode(pdf_base64)
        except Exception as e:
            return jsonify({
                'error': 'Invalid Base64',
                'message': f'Failed to decode base64: {str(e)}'
            }), 400
        
        # Create file object
        pdf_file = io.BytesIO(pdf_bytes)
        
        # Try pdfplumber first (handles tables + layout)
        full_text = ''
        page_texts = []
        extraction_method = 'pdfplumber'
        
        try:
            with pdfplumber.open(pdf_file) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    # Extract text with layout
                    page_text = page.extract_text() or ""
                    
                    # Extract tables
                    table_text = extract_tables_from_page(page)
                    
                    # Combine
                    combined_text = page_text + table_text
                    
                    page_texts.append({
                        'page_number': page_num + 1,
                        'text': combined_text,
                        'char_count': len(combined_text),
                        'has_tables': bool(table_text),
                        'method': 'pdfplumber'
                    })
                    
                    full_text += combined_text + '\n'
            
            # Check if text is empty (might be scanned PDF)
            if len(full_text.strip()) < 100:
                # Fallback to OCR
                extraction_method = 'OCR_fallback'
                pdf_file.seek(0)  # Reset file pointer
                full_text, page_texts = extract_with_ocr(pdf_bytes)
        
        except Exception as e:
            # If pdfplumber fails, try OCR
            extraction_method = 'OCR_fallback'
            full_text, page_texts = extract_with_ocr(pdf_bytes)
        
        # Clean up extra whitespace
        full_text = ' '.join(full_text.split())
        
        return jsonify({
            'success': True,
            'text': full_text,
            'total_pages': len(page_texts),
            'total_chars': len(full_text),
            'pages': page_texts,
            'extraction_method': extraction_method,
            'metadata': {
                'file_size_bytes': len(pdf_bytes),
                'has_tables': any(p.get('has_tables') for p in page_texts)
            }
        }), 200
    
    except Exception as e:
        return jsonify({
            'error': 'Extraction Failed',
            'message': str(e)
        }), 500

# =============================
# RUN SERVER
# =============================

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)