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
import openpyxl
import xlrd
from datetime import datetime

app = Flask(__name__)
CORS(app)

# =============================
# SECURITY - API KEY
# =============================

API_KEY = os.environ.get('API_KEY', 'fp_document_service_2026_key')

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
        'service': 'Firepulse Document Extraction Service',
        'version': '2.0.0',
        'capabilities': [
            'PDF text extraction with layout preservation',
            'PDF table structure extraction',
            'OCR for scanned PDFs',
            'Excel (.xls, .xlsx) multi-sheet extraction',
            'Excel formula and value extraction',
            'Excel merged cell handling'
        ],
        'endpoints': {
            'POST /extract-pdf': 'Extract text from PDF files',
            'POST /extract-excel': 'Extract text from Excel files (.xls, .xlsx)'
        }
    })

# =============================
# HELPER: OCR FOR SCANNED PDFs
# =============================

def extract_with_ocr(pdf_bytes):
    """Extract text from scanned PDF using OCR"""
    try:
        images = convert_from_bytes(pdf_bytes)
        
        full_text = ''
        page_texts = []
        
        for page_num, image in enumerate(images):
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
# HELPER: EXTRACT TABLES FROM PDF
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
            row_text = " | ".join([str(cell) if cell else "" for cell in row])
            table_text += row_text + "\n"
        
        table_text += "[END TABLE]\n"
    
    return table_text

# =============================
# HELPER: FORMAT EXCEL CELL VALUE
# =============================

def format_excel_value(cell, include_formula=False):
    """Format Excel cell value with type handling"""
    if cell is None:
        return ''
    
    # Handle different cell types
    if isinstance(cell, (int, float)):
        return str(cell)
    elif isinstance(cell, datetime):
        return cell.strftime('%Y-%m-%d %H:%M:%S')
    elif isinstance(cell, bool):
        return 'TRUE' if cell else 'FALSE'
    else:
        return str(cell).strip()

# =============================
# HELPER: EXTRACT XLSX (OpenPyXL)
# =============================

def extract_xlsx(file_bytes):
    """Extract all sheets from .xlsx file"""
    try:
        workbook = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        sheets_data = []
        full_text = ''
        
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            sheet_text = f"\n{'='*50}\n[SHEET: {sheet_name}]\n{'='*50}\n"
            
            # Get dimensions
            max_row = sheet.max_row
            max_col = sheet.max_column
            
            if max_row == 0 or max_col == 0:
                sheet_text += "(Empty sheet)\n"
                continue
            
            # Extract as table
            rows_data = []
            for row in sheet.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
                row_values = []
                for cell in row:
                    # Get cell value
                    value = format_excel_value(cell.value)
                    
                    # Optionally get formula if exists
                    if cell.data_type == 'f' and cell.value:
                        formula = str(cell.value) if hasattr(cell, 'value') else ''
                        value = f"{value} (Formula: {formula})"
                    
                    row_values.append(value)
                
                rows_data.append(row_values)
            
            # Format as table
            if rows_data:
                # Assume first row is header
                headers = rows_data[0]
                sheet_text += "Columns: " + " | ".join(headers) + "\n"
                sheet_text += "-" * 80 + "\n"
                
                # Add data rows
                for row_idx, row in enumerate(rows_data[1:], start=2):
                    row_text = " | ".join(row)
                    sheet_text += f"Row {row_idx}: {row_text}\n"
            
            sheet_text += f"\n(Sheet has {max_row} rows, {max_col} columns)\n"
            
            sheets_data.append({
                'sheet_name': sheet_name,
                'text': sheet_text,
                'rows': max_row,
                'cols': max_col
            })
            
            full_text += sheet_text + "\n"
        
        return {
            'full_text': full_text,
            'sheets': sheets_data,
            'total_sheets': len(sheets_data)
        }
    
    except Exception as e:
        raise Exception(f'XLSX extraction failed: {str(e)}')

# =============================
# HELPER: EXTRACT XLS (xlrd)
# =============================

def extract_xls(file_bytes):
    """Extract all sheets from .xls file"""
    try:
        workbook = xlrd.open_workbook(file_contents=file_bytes, formatting_info=False)
        sheets_data = []
        full_text = ''
        
        for sheet_idx in range(workbook.nsheets):
            sheet = workbook.sheet_by_index(sheet_idx)
            sheet_name = sheet.name
            sheet_text = f"\n{'='*50}\n[SHEET: {sheet_name}]\n{'='*50}\n"
            
            if sheet.nrows == 0 or sheet.ncols == 0:
                sheet_text += "(Empty sheet)\n"
                continue
            
            # Extract as table
            rows_data = []
            for row_idx in range(sheet.nrows):
                row_values = []
                for col_idx in range(sheet.ncols):
                    cell = sheet.cell(row_idx, col_idx)
                    
                    # Handle different cell types
                    if cell.ctype == xlrd.XL_CELL_EMPTY:
                        value = ''
                    elif cell.ctype == xlrd.XL_CELL_TEXT:
                        value = str(cell.value)
                    elif cell.ctype == xlrd.XL_CELL_NUMBER:
                        value = str(cell.value)
                    elif cell.ctype == xlrd.XL_CELL_DATE:
                        date_tuple = xlrd.xldate_as_tuple(cell.value, workbook.datemode)
                        value = f"{date_tuple[0]}-{date_tuple[1]:02d}-{date_tuple[2]:02d}"
                    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                        value = 'TRUE' if cell.value else 'FALSE'
                    else:
                        value = str(cell.value)
                    
                    row_values.append(value)
                
                rows_data.append(row_values)
            
            # Format as table
            if rows_data:
                # First row as header
                headers = rows_data[0]
                sheet_text += "Columns: " + " | ".join(headers) + "\n"
                sheet_text += "-" * 80 + "\n"
                
                # Data rows
                for row_idx, row in enumerate(rows_data[1:], start=2):
                    row_text = " | ".join(row)
                    sheet_text += f"Row {row_idx}: {row_text}\n"
            
            sheet_text += f"\n(Sheet has {sheet.nrows} rows, {sheet.ncols} columns)\n"
            
            sheets_data.append({
                'sheet_name': sheet_name,
                'text': sheet_text,
                'rows': sheet.nrows,
                'cols': sheet.ncols
            })
            
            full_text += sheet_text + "\n"
        
        return {
            'full_text': full_text,
            'sheets': sheets_data,
            'total_sheets': len(sheets_data)
        }
    
    except Exception as e:
        raise Exception(f'XLS extraction failed: {str(e)}')

# =============================
# ENDPOINT: EXTRACT PDF
# =============================

@app.route('/extract-pdf', methods=['POST'])
@require_api_key
def extract_pdf():
    try:
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
        
        # Try pdfplumber first
        full_text = ''
        page_texts = []
        extraction_method = 'pdfplumber'
        
        try:
            with pdfplumber.open(pdf_file) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    page_text = page.extract_text() or ""
                    table_text = extract_tables_from_page(page)
                    combined_text = page_text + table_text
                    
                    page_texts.append({
                        'page_number': page_num + 1,
                        'text': combined_text,
                        'char_count': len(combined_text),
                        'has_tables': bool(table_text),
                        'method': 'pdfplumber'
                    })
                    
                    full_text += combined_text + '\n'
            
            # Check if text is empty (scanned PDF)
            if len(full_text.strip()) < 100:
                extraction_method = 'OCR_fallback'
                pdf_file.seek(0)
                full_text, page_texts = extract_with_ocr(pdf_bytes)
        
        except Exception as e:
            # Fallback to OCR
            extraction_method = 'OCR_fallback'
            full_text, page_texts = extract_with_ocr(pdf_bytes)
        
        # Clean whitespace
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
            'error': 'PDF Extraction Failed',
            'message': str(e)
        }), 500

# =============================
# ENDPOINT: EXTRACT EXCEL
# =============================

@app.route('/extract-excel', methods=['POST'])
@require_api_key
def extract_excel():
    try:
        data = request.json
        
        if not data or 'file_base64' not in data:
            return jsonify({
                'error': 'Bad Request',
                'message': 'Missing file_base64 in request body'
            }), 400
        
        # Get file type
        file_extension = data.get('file_extension', '.xlsx').lower()
        
        # Decode base64
        file_base64 = data['file_base64']
        
        try:
            file_bytes = base64.b64decode(file_base64)
        except Exception as e:
            return jsonify({
                'error': 'Invalid Base64',
                'message': f'Failed to decode base64: {str(e)}'
            }), 400
        
        # Extract based on file type
        if file_extension == '.xls':
            result = extract_xls(file_bytes)
        else:  # .xlsx
            result = extract_xlsx(file_bytes)
        
        return jsonify({
            'success': True,
            'text': result['full_text'],
            'total_sheets': result['total_sheets'],
            'total_chars': len(result['full_text']),
            'sheets': result['sheets'],
            'extraction_method': 'openpyxl' if file_extension == '.xlsx' else 'xlrd',
            'metadata': {
                'file_size_bytes': len(file_bytes),
                'file_extension': file_extension
            }
        }), 200
    
    except Exception as e:
        return jsonify({
            'error': 'Excel Extraction Failed',
            'message': str(e)
        }), 500

# =============================
# RUN SERVER
# =============================

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)