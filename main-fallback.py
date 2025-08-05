import os
import re
import json
import zipfile
import tempfile
import logging
from difflib import get_close_matches
from werkzeug.utils import secure_filename
from flask import Flask, render_template_string, request, jsonify, send_file, session
import pandas as pd
try:
    import PyPDF2 as pdf_lib
    PDF_LIB = 'pypdf2'
except ImportError:
    try:
        import fitz
        PDF_LIB = 'pymupdf'
    except ImportError:
        PDF_LIB = None

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'change-this-in-production')

UPLOAD_FOLDER = 'uploads'
RESULTS_FOLDER = 'results'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULTS_FOLDER, exist_ok=True)

US_STATES = ['AL','AK','AZ','AR','CA','CO','CT','DE','FL','GA','HI','ID','IL','IN','IA','KS','KY','LA','ME','MD','MA','MI','MN','MS','MO','MT','NE','NV','NH','NJ','NM','NY','NC','ND','OH','OK','OR','PA','RI','SC','SD','TN','TX','UT','VT','VA','WA','WV','WI','WY','DC']

logging.basicConfig(level=logging.INFO)

def safe_filename(filename):
    return secure_filename(filename)

def validate_pdf_pypdf2(file_path):
    try:
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            if len(reader.pages) > 0:
                first_page = reader.pages[0]
                text = first_page.extract_text()
                return len(text.strip()) > 0
        return False
    except:
        return False

def validate_pdf_pymupdf(file_path):
    try:
        doc = fitz.open(file_path)
        has_text = any(page.get_text().strip() for page in doc)
        doc.close()
        return has_text
    except:
        return False

def validate_pdf(file_path):
    if PDF_LIB == 'pypdf2':
        return validate_pdf_pypdf2(file_path)
    elif PDF_LIB == 'pymupdf':
        return validate_pdf_pymupdf(file_path)
    else:
        return False

@app.route('/')
def index():
    return render_template_string(HOME_TEMPLATE)

@app.route('/api/statement-separator', methods=['POST'])
def statement_separator():
    try:
        pdf_file = request.files.get('pdf_file')
        excel_file = request.files.get('excel_file')
        
        if not pdf_file or not excel_file:
            return jsonify({'error': 'Both PDF and Excel files required'}), 400
            
        temp_dir = tempfile.mkdtemp()
        pdf_path = os.path.join(temp_dir, safe_filename(pdf_file.filename))
        excel_path = os.path.join(temp_dir, safe_filename(excel_file.filename))
        
        pdf_file.save(pdf_path)
        excel_file.save(excel_path)
        
        if not validate_pdf(pdf_path):
            return jsonify({'error': 'PDF contains no readable text'}), 400
            
        result = process_statements(pdf_path, excel_path)
        
        if result.get('pending_decisions'):
            session['temp_data'] = {
                'pdf_path': pdf_path,
                'categories': result['categories'],
                'pending': result['pending_decisions']
            }
            return jsonify({'status': 'review', 'decisions': result['pending_decisions']})
        else:
            files = create_statement_pdfs(pdf_path, result['categories'])
            return jsonify({'status': 'complete', 'files': files})
            
    except Exception as e:
        logging.error(f"Statement separator error: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/statement-decision', methods=['POST'])
def statement_decision():
    try:
        data = request.json
        action = data.get('action')
        statement = data.get('statement')
        
        temp_data = session.get('temp_data')
        if not temp_data:
            return jsonify({'error': 'Session expired'}), 400
            
        categories = temp_data['categories']
        pending = temp_data['pending']
        
        pages = list(range(statement['page_num'], statement['page_num'] + statement['total_pages']))
        
        if action == 'dnm':
            categories['dnm'].extend(pages)
        elif action == 'foreign':
            categories['foreign'].extend(pages)
        elif action == 'national':
            if statement['total_pages'] == 1:
                categories['national_single'].extend(pages)
            else:
                categories['national_multi'].extend(pages)
                
        pending.remove(statement)
        
        if not pending:
            files = create_statement_pdfs(temp_data['pdf_path'], categories)
            return jsonify({'status': 'complete', 'files': files})
        else:
            session['temp_data'] = temp_data
            return jsonify({'status': 'continue', 'remaining': len(pending)})
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/invoice-processor', methods=['POST'])
def invoice_processor():
    try:
        pdf_file = request.files.get('pdf_file')
        if not pdf_file:
            return jsonify({'error': 'PDF file required'}), 400
            
        temp_dir = tempfile.mkdtemp()
        pdf_path = os.path.join(temp_dir, safe_filename(pdf_file.filename))
        pdf_file.save(pdf_path)
        
        invoices = extract_invoices(pdf_path)
        if not invoices:
            return jsonify({'error': 'No invoices found in PDF'}), 400
            
        zip_path = create_invoice_zip(pdf_path, invoices)
        return send_file(zip_path, as_attachment=True, download_name='invoices.zip')
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/excel-processor', methods=['POST'])
def excel_processor():
    try:
        excel_file = request.files.get('excel_file')
        if not excel_file:
            return jsonify({'error': 'Excel file required'}), 400
            
        temp_path = os.path.join(tempfile.gettempdir(), safe_filename(excel_file.filename))
        excel_file.save(temp_path)
        
        df = pd.read_excel(temp_path)
        data = []
        for _, row in df.iterrows():
            data.append({
                'd': str(row.iloc[0]) if len(row) > 0 else '',
                's': str(row.iloc[1]) if len(row) > 1 else '',
                'w': str(row.iloc[2]) if len(row) > 2 else ''
            })
            
        os.remove(temp_path)
        return jsonify({'status': 'success', 'data': data})
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(RESULTS_FOLDER, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return jsonify({'error': 'File not found'}), 404

def process_statements_pypdf2(pdf_path, excel_path):
    df = pd.read_excel(excel_path)
    company_names = df.iloc[:, 0].dropna().tolist()
    
    categories = {'dnm': [], 'national_single': [], 'national_multi': [], 'foreign': []}
    pending_decisions = []
    
    start_markers = ["914.949.9618", "302.703.8961", "www.unitedcorporate.com", "AR@UNITEDCORPORATE.COM"]
    end_marker = "STATEMENT OF OPEN INVOICE(S)"
    
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        
        for page_num, page in enumerate(reader.pages):
            text = page.extract_text()
            
            page_match = re.search(r'Page\s*(\d+)\s*of\s*(\d+)', text, re.IGNORECASE)
            total_pages = int(page_match.group(2)) if page_match else 1
            
            start_idx = min((text.find(marker) for marker in start_markers if text.find(marker) != -1), default=-1)
            end_idx = text.find(end_marker)
            
            if start_idx != -1 and end_idx != -1 and start_idx < end_idx:
                extracted_text = text[start_idx:end_idx].strip()
                lines = [line.strip() for line in extracted_text.splitlines() if line.strip()]
                
                if lines:
                    company_name = lines[0]
                    pages = [page_num + 1]  # PyPDF2 processes page by page
                    
                    if company_name in company_names:
                        categories['dnm'].extend(pages)
                    else:
                        close_matches = get_close_matches(company_name, company_names, n=1, cutoff=0.8)
                        if close_matches:
                            pending_decisions.append({
                                'page_num': page_num + 1,
                                'company_name': company_name,
                                'close_match': close_matches[0],
                                'total_pages': 1,
                                'lines': lines[:5]
                            })
                        else:
                            if "email" in text.lower():
                                categories['dnm'].extend(pages)
                            elif any(state in ' '.join(lines[1:]) for state in US_STATES):
                                categories['national_single'].extend(pages)
                            else:
                                categories['foreign'].extend(pages)
                                
    return {'categories': categories, 'pending_decisions': pending_decisions}

def process_statements_pymupdf(pdf_path, excel_path):
    df = pd.read_excel(excel_path)
    company_names = df.iloc[:, 0].dropna().tolist()
    
    doc = fitz.open(pdf_path)
    categories = {'dnm': [], 'national_single': [], 'national_multi': [], 'foreign': []}
    pending_decisions = []
    
    start_markers = ["914.949.9618", "302.703.8961", "www.unitedcorporate.com", "AR@UNITEDCORPORATE.COM"]
    end_marker = "STATEMENT OF OPEN INVOICE(S)"
    
    page_num = 0
    while page_num < len(doc):
        page = doc.load_page(page_num)
        text = page.get_text()
        
        page_match = re.search(r'Page\s*(\d+)\s*of\s*(\d+)', text, re.IGNORECASE)
        total_pages = int(page_match.group(2)) if page_match else 1
        
        start_idx = min((text.find(marker) for marker in start_markers if text.find(marker) != -1), default=-1)
        end_idx = text.find(end_marker)
        
        if start_idx != -1 and end_idx != -1 and start_idx < end_idx:
            extracted_text = text[start_idx:end_idx].strip()
            lines = [line.strip() for line in extracted_text.splitlines() if line.strip()]
            
            if lines:
                company_name = lines[0]
                pages = list(range(page_num + 1, page_num + 1 + total_pages))
                
                if company_name in company_names:
                    categories['dnm'].extend(pages)
                else:
                    close_matches = get_close_matches(company_name, company_names, n=1, cutoff=0.8)
                    if close_matches:
                        pending_decisions.append({
                            'page_num': page_num + 1,
                            'company_name': company_name,
                            'close_match': close_matches[0],
                            'total_pages': total_pages,
                            'lines': lines[:5]
                        })
                    else:
                        if "email" in text.lower():
                            categories['dnm'].extend(pages)
                        elif any(state in ' '.join(lines[1:]) for state in US_STATES):
                            if total_pages == 1:
                                categories['national_single'].extend(pages)
                            else:
                                categories['national_multi'].extend(pages)
                        else:
                            categories['foreign'].extend(pages)
                            
        page_num += total_pages if page_match else 1
        
    doc.close()
    return {'categories': categories, 'pending_decisions': pending_decisions}

def process_statements(pdf_path, excel_path):
    if PDF_LIB == 'pypdf2':
        return process_statements_pypdf2(pdf_path, excel_path)
    elif PDF_LIB == 'pymupdf':
        return process_statements_pymupdf(pdf_path, excel_path)
    else:
        raise Exception("No PDF library available")

def create_statement_pdfs_pypdf2(pdf_path, categories):
    files = []
    
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        
        for name, pages in categories.items():
            if pages:
                writer = PyPDF2.PdfWriter()
                for page_num in sorted(set(pages)):
                    if 0 <= page_num - 1 < len(reader.pages):
                        writer.add_page(reader.pages[page_num - 1])
                        
                if len(writer.pages) > 0:
                    filename = f"{name}.pdf"
                    file_path = os.path.join(RESULTS_FOLDER, filename)
                    with open(file_path, 'wb') as output_file:
                        writer.write(output_file)
                    files.append(filename)
                    
    return files

def create_statement_pdfs_pymupdf(pdf_path, categories):
    doc = fitz.open(pdf_path)
    files = []
    
    for name, pages in categories.items():
        if pages:
            output_doc = fitz.open()
            for page_num in sorted(set(pages)):
                if 0 <= page_num - 1 < len(doc):
                    output_doc.insert_pdf(doc, from_page=page_num-1, to_page=page_num-1)
                    
            if output_doc.page_count > 0:
                filename = f"{name}.pdf"
                file_path = os.path.join(RESULTS_FOLDER, filename)
                output_doc.save(file_path)
                files.append(filename)
            output_doc.close()
            
    doc.close()
    return files

def create_statement_pdfs(pdf_path, categories):
    if PDF_LIB == 'pypdf2':
        return create_statement_pdfs_pypdf2(pdf_path, categories)
    elif PDF_LIB == 'pymupdf':
        return create_statement_pdfs_pymupdf(pdf_path, categories)
    else:
        raise Exception("No PDF library available")

def extract_invoices_pypdf2(pdf_path):
    pattern = r'\b[PR]\d{6,8}\b'
    invoices = {}
    
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        
        for page_num, page in enumerate(reader.pages):
            text = page.extract_text()
            invoice_numbers = re.findall(pattern, text)
            
            for invoice_number in invoice_numbers:
                if invoice_number not in invoices:
                    invoices[invoice_number] = []
                invoices[invoice_number].append(page_num)
                
    return invoices

def extract_invoices_pymupdf(pdf_path):
    doc = fitz.open(pdf_path)
    pattern = r'\b[PR]\d{6,8}\b'
    invoices = {}
    
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        invoice_numbers = re.findall(pattern, text)
        
        for invoice_number in invoice_numbers:
            if invoice_number not in invoices:
                invoices[invoice_number] = []
            invoices[invoice_number].append(page_num)
            
    doc.close()
    return invoices

def extract_invoices(pdf_path):
    if PDF_LIB == 'pypdf2':
        return extract_invoices_pypdf2(pdf_path)
    elif PDF_LIB == 'pymupdf':
        return extract_invoices_pymupdf(pdf_path)
    else:
        raise Exception("No PDF library available")

def create_invoice_zip_pypdf2(pdf_path, invoices):
    zip_path = os.path.join(RESULTS_FOLDER, 'invoices.zip')
    
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for invoice_number, pages in invoices.items():
                writer = PyPDF2.PdfWriter()
                for page_num in pages:
                    writer.add_page(reader.pages[page_num])
                    
                if len(writer.pages) > 0:
                    temp_path = f"/tmp/{invoice_number}.pdf"
                    with open(temp_path, 'wb') as temp_file:
                        writer.write(temp_file)
                    zipf.write(temp_path, f"{invoice_number}.pdf")
                    os.remove(temp_path)
                    
    return zip_path

def create_invoice_zip_pymupdf(pdf_path, invoices):
    doc = fitz.open(pdf_path)
    zip_path = os.path.join(RESULTS_FOLDER, 'invoices.zip')
    
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for invoice_number, pages in invoices.items():
            output_doc = fitz.open()
            for page_num in pages:
                output_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
                
            if output_doc.page_count > 0:
                temp_path = f"/tmp/{invoice_number}.pdf"
                output_doc.save(temp_path)
                zipf.write(temp_path, f"{invoice_number}.pdf")
                os.remove(temp_path)
            output_doc.close()
            
    doc.close()
    return zip_path

def create_invoice_zip(pdf_path, invoices):
    if PDF_LIB == 'pypdf2':
        return create_invoice_zip_pypdf2(pdf_path, invoices)
    elif PDF_LIB == 'pymupdf':
        return create_invoice_zip_pymupdf(pdf_path, invoices)
    else:
        raise Exception("No PDF library available")

HOME_TEMPLATE = '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document Processor</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh; display: flex; align-items: center; justify-content: center;
        }
        .container { 
            background: white; border-radius: 20px; padding: 40px; 
            box-shadow: 0 20px 40px rgba(0,0,0,0.1); max-width: 800px; width: 90%;
        }
        h1 { text-align: center; color: #333; margin-bottom: 40px; font-size: 2.5rem; }
        .tools { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; }
        .tool { 
            border: 2px solid #e0e0e0; border-radius: 15px; padding: 25px; 
            transition: all 0.3s ease; cursor: pointer; background: #f9f9f9;
        }
        .tool:hover { border-color: #667eea; transform: translateY(-5px); box-shadow: 0 10px 25px rgba(0,0,0,0.1); }
        .tool h3 { color: #333; margin-bottom: 15px; font-size: 1.3rem; }
        .tool p { color: #666; margin-bottom: 20px; line-height: 1.5; }
        .upload-area { 
            border: 2px dashed #ccc; border-radius: 10px; padding: 20px; 
            text-align: center; margin-bottom: 15px; transition: all 0.3s ease;
        }
        .upload-area.dragover { border-color: #667eea; background: #f0f4ff; }
        .btn { 
            background: #667eea; color: white; border: none; padding: 12px 24px; 
            border-radius: 8px; cursor: pointer; font-size: 1rem; transition: all 0.3s ease;
        }
        .btn:hover { background: #5a6fd8; transform: translateY(-2px); }
        .btn:disabled { background: #ccc; cursor: not-allowed; transform: none; }
        .result { margin-top: 20px; padding: 15px; border-radius: 8px; }
        .success { background: #d4edda; color: #155724; border: 1px solid #c3e6cb; }
        .error { background: #f8d7da; color: #721c24; border: 1px solid #f5c6cb; }
        .loading { text-align: center; color: #667eea; }
        .hidden { display: none; }
        .file-info { font-size: 0.9rem; color: #666; margin-top: 10px; }
        .decision-panel { background: #fff3cd; border: 1px solid #ffeaa7; border-radius: 8px; padding: 20px; margin: 15px 0; }
        .decision-buttons { display: flex; gap: 10px; margin-top: 15px; }
        .decision-buttons button { padding: 8px 16px; border: none; border-radius: 5px; cursor: pointer; }
        .btn-dnm { background: #dc3545; color: white; }
        .btn-national { background: #28a745; color: white; }
        .btn-foreign { background: #ffc107; color: #212529; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Document Processor</h1>
        <div class="tools">
            <!-- Statement Separator -->
            <div class="tool">
                <h3>📋 Statement Separator</h3>
                <p>Upload PDF statements and Excel DNM list to automatically categorize by location and mailing preferences.</p>
                <div class="upload-area" id="stmt-upload">
                    <p>Drop files here or click to select</p>
                    <input type="file" id="stmt-pdf" accept=".pdf" style="display:none">
                    <input type="file" id="stmt-excel" accept=".xlsx,.xls" style="display:none">
                    <div class="file-info" id="stmt-files"></div>
                </div>
                <button class="btn" onclick="processStatements()" id="stmt-btn" disabled>Process Statements</button>
                <div id="stmt-result" class="result hidden"></div>
                <div id="stmt-decisions" class="hidden"></div>
            </div>

            <!-- Invoice Processor -->
            <div class="tool">
                <h3>🧾 Invoice Processor</h3>
                <p>Extract and separate individual invoices from a combined PDF file using pattern recognition.</p>
                <div class="upload-area" id="inv-upload">
                    <p>Drop PDF file here or click to select</p>
                    <input type="file" id="inv-pdf" accept=".pdf" style="display:none">
                    <div class="file-info" id="inv-files"></div>
                </div>
                <button class="btn" onclick="processInvoices()" id="inv-btn" disabled>Extract Invoices</button>
                <div id="inv-result" class="result hidden"></div>
            </div>

            <!-- Excel Processor -->
            <div class="tool">
                <h3>📊 Excel Processor</h3>
                <p>Convert Excel data into structured format for batch processing and form automation.</p>
                <div class="upload-area" id="excel-upload">
                    <p>Drop Excel file here or click to select</p>
                    <input type="file" id="excel-file" accept=".xlsx,.xls" style="display:none">
                    <div class="file-info" id="excel-files"></div>
                </div>
                <button class="btn" onclick="processExcel()" id="excel-btn" disabled>Process Excel</button>
                <div id="excel-result" class="result hidden"></div>
            </div>
        </div>
    </div>

    <script>
        // File upload handlers
        function setupUpload(uploadId, inputIds, btnId, fileInfoId) {
            const uploadArea = document.getElementById(uploadId);
            const btn = document.getElementById(btnId);
            const fileInfo = document.getElementById(fileInfoId);
            
            uploadArea.addEventListener('click', () => {
                if (inputIds.length === 1) {
                    document.getElementById(inputIds[0]).click();
                } else {
                    // For multiple files, cycle through them
                    let currentInput = 0;
                    const clickNext = () => {
                        if (currentInput < inputIds.length) {
                            document.getElementById(inputIds[currentInput]).click();
                            currentInput++;
                        }
                    };
                    clickNext();
                }
            });
            
            inputIds.forEach(inputId => {
                const input = document.getElementById(inputId);
                input.addEventListener('change', () => updateFileInfo(inputIds, fileInfo, btn));
            });
            
            // Drag and drop
            ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
                uploadArea.addEventListener(eventName, preventDefaults);
            });
            
            ['dragenter', 'dragover'].forEach(eventName => {
                uploadArea.addEventListener(eventName, () => uploadArea.classList.add('dragover'));
            });
            
            ['dragleave', 'drop'].forEach(eventName => {
                uploadArea.addEventListener(eventName, () => uploadArea.classList.remove('dragover'));
            });
            
            uploadArea.addEventListener('drop', (e) => {
                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    document.getElementById(inputIds[0]).files = files;
                    updateFileInfo(inputIds, fileInfo, btn);
                }
            });
        }
        
        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }
        
        function updateFileInfo(inputIds, fileInfo, btn) {
            const files = inputIds.map(id => {
                const input = document.getElementById(id);
                return input.files[0] ? input.files[0].name : null;
            }).filter(Boolean);
            
            if (inputIds.length === 1 && files.length > 0) {
                fileInfo.textContent = files[0];
                btn.disabled = false;
            } else if (inputIds.length === 2 && files.length === 2) {
                fileInfo.textContent = files.join(', ');
                btn.disabled = false;
            } else {
                fileInfo.textContent = files.join(', ') + (inputIds.length === 2 && files.length === 1 ? ' (need both PDF and Excel)' : '');
                btn.disabled = inputIds.length === 2 ? files.length !== 2 : files.length === 0;
            }
        }
        
        // Initialize uploads
        setupUpload('stmt-upload', ['stmt-pdf', 'stmt-excel'], 'stmt-btn', 'stmt-files');
        setupUpload('inv-upload', ['inv-pdf'], 'inv-btn', 'inv-files');
        setupUpload('excel-upload', ['excel-file'], 'excel-btn', 'excel-files');
        
        // Processing functions
        async function processStatements() {
            const btn = document.getElementById('stmt-btn');
            const result = document.getElementById('stmt-result');
            const decisions = document.getElementById('stmt-decisions');
            
            btn.disabled = true;
            btn.textContent = 'Processing...';
            result.className = 'result loading';
            result.innerHTML = 'Processing statements...';
            result.classList.remove('hidden');
            
            const formData = new FormData();
            formData.append('pdf_file', document.getElementById('stmt-pdf').files[0]);
            formData.append('excel_file', document.getElementById('stmt-excel').files[0]);
            
            try {
                const response = await fetch('/api/statement-separator', {
                    method: 'POST',
                    body: formData
                });
                
                const data = await response.json();
                
                if (data.status === 'review') {
                    result.className = 'result';
                    result.innerHTML = `<p>Found ${data.decisions.length} statements requiring review.</p>`;
                    showDecisions(data.decisions);
                } else if (data.status === 'complete') {
                    result.className = 'result success';
                    result.innerHTML = `<p>✅ Processing complete! Generated ${data.files.length} PDF files.</p>
                        ${data.files.map(file => `<a href="/download/${file}" class="btn" style="margin: 5px;">${file}</a>`).join('')}`;
                } else {
                    throw new Error(data.error || 'Processing failed');
                }
            } catch (error) {
                result.className = 'result error';
                result.innerHTML = `<p>❌ Error: ${error.message}</p>`;
            }
            
            btn.disabled = false;
            btn.textContent = 'Process Statements';
        }
        
        function showDecisions(decisionsData) {
            const container = document.getElementById('stmt-decisions');
            container.innerHTML = '';
            
            decisionsData.forEach((decision, index) => {
                const panel = document.createElement('div');
                panel.className = 'decision-panel';
                panel.innerHTML = `
                    <h4>Statement ${index + 1}: "${decision.company_name}"</h4>
                    <p><strong>Suggested match:</strong> ${decision.close_match}</p>
                    <p><strong>Pages:</strong> ${decision.total_pages}</p>
                    <div class="decision-buttons">
                        <button class="btn-dnm" onclick="makeDecision(${index}, 'dnm')">DNM List</button>
                        <button class="btn-national" onclick="makeDecision(${index}, 'national')">National</button>
                        <button class="btn-foreign" onclick="makeDecision(${index}, 'foreign')">Foreign</button>
                    </div>
                `;
                container.appendChild(panel);
            });
            
            container.classList.remove('hidden');
            window.currentDecisions = decisionsData;
        }
        
        async function makeDecision(index, action) {
            const decision = window.currentDecisions[index];
            
            try {
                const response = await fetch('/api/statement-decision', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ action, statement: decision })
                });
                
                const data = await response.json();
                
                // Remove the decided panel
                document.getElementsByClassName('decision-panel')[index].remove();
                window.currentDecisions.splice(index, 1);
                
                if (data.status === 'complete') {
                    document.getElementById('stmt-decisions').classList.add('hidden');
                    const result = document.getElementById('stmt-result');
                    result.className = 'result success';
                    result.innerHTML = `<p>✅ All decisions processed! Generated PDF files.</p>
                        ${data.files.map(file => `<a href="/download/${file}" class="btn" style="margin: 5px;">${file}</a>`).join('')}`;
                }
            } catch (error) {
                alert('Error processing decision: ' + error.message);
            }
        }
        
        async function processInvoices() {
            const btn = document.getElementById('inv-btn');
            const result = document.getElementById('inv-result');
            
            btn.disabled = true;
            btn.textContent = 'Processing...';
            result.className = 'result loading';
            result.innerHTML = 'Extracting invoices...';
            result.classList.remove('hidden');
            
            const formData = new FormData();
            formData.append('pdf_file', document.getElementById('inv-pdf').files[0]);
            
            try {
                const response = await fetch('/api/invoice-processor', {
                    method: 'POST',
                    body: formData
                });
                
                if (response.ok) {
                    const blob = await response.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = 'invoices.zip';
                    a.click();
                    
                    result.className = 'result success';
                    result.innerHTML = '<p>✅ Invoices extracted successfully! Download started.</p>';
                } else {
                    const data = await response.json();
                    throw new Error(data.error || 'Processing failed');
                }
            } catch (error) {
                result.className = 'result error';
                result.innerHTML = `<p>❌ Error: ${error.message}</p>`;
            }
            
            btn.disabled = false;
            btn.textContent = 'Extract Invoices';
        }
        
        async function processExcel() {
            const btn = document.getElementById('excel-btn');
            const result = document.getElementById('excel-result');
            
            btn.disabled = true;
            btn.textContent = 'Processing...';
            result.className = 'result loading';
            result.innerHTML = 'Processing Excel data...';
            result.classList.remove('hidden');
            
            const formData = new FormData();
            formData.append('excel_file', document.getElementById('excel-file').files[0]);
            
            try {
                const response = await fetch('/api/excel-processor', {
                    method: 'POST',
                    body: formData
                });
                
                const data = await response.json();
                
                if (data.status === 'success') {
                    result.className = 'result success';
                    result.innerHTML = `
                        <p>✅ Processed ${data.data.length} rows successfully!</p>
                        <details style="margin-top: 10px;">
                            <summary>View processed data</summary>
                            <pre style="background: #f5f5f5; padding: 10px; margin-top: 10px; border-radius: 4px; overflow-x: auto;">
${JSON.stringify(data.data.slice(0, 5), null, 2)}${data.data.length > 5 ? '\\n... and ' + (data.data.length - 5) + ' more rows' : ''}
                            </pre>
                        </details>
                    `;
                } else {
                    throw new Error(data.error || 'Processing failed');
                }
            } catch (error) {
                result.className = 'result error';
                result.innerHTML = `<p>❌ Error: ${error.message}</p>`;
            }
            
            btn.disabled = false;
            btn.textContent = 'Process Excel';
        }
    </script>
</body>
</html>
'''

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))