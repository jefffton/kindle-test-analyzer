import os
import re
import json
import csv
import io
import threading
import uuid
import time
import chardet
from flask import Flask, request, render_template, jsonify
from werkzeug.utils import secure_filename
import anthropic

# Document parsers
from docx import Document
from pdfminer.high_level import extract_text as pdf_extract_text
from openpyxl import load_workbook
from bs4 import BeautifulSoup

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')

ALLOWED_EXTENSIONS = {
    'txt', 'doc', 'docx', 'pdf', 'xlsx', 'xls', 'csv',
    'html', 'htm', 'md', 'json', 'xml', 'rtf'
}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# In-memory job store: {job_id: {status, result, error, created_at}}
_jobs = {}
_jobs_lock = threading.Lock()


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_text_from_file(filepath, filename):
    ext = filename.rsplit('.', 1)[1].lower()

    if ext == 'pdf':
        return pdf_extract_text(filepath)

    if ext in ('docx', 'doc'):
        doc = Document(filepath)
        parts = []
        for para in doc.paragraphs:
            if para.text.strip():
                parts.append(para.text)
        for table in doc.tables:
            for row in table.rows:
                row_text = ' | '.join(cell.text.strip() for cell in row.cells if cell.text.strip())
                if row_text:
                    parts.append(row_text)
        return '\n'.join(parts)

    if ext == 'xlsx':
        wb = load_workbook(filepath, read_only=True, data_only=True)
        parts = []
        for sheet in wb.worksheets:
            parts.append(f'[Sheet: {sheet.title}]')
            for row in sheet.iter_rows(values_only=True):
                row_text = ' | '.join(str(c) for c in row if c is not None)
                if row_text.strip():
                    parts.append(row_text)
        return '\n'.join(parts)

    if ext == 'xls':
        import xlrd

        # Try xlrd first (true legacy .xls binary format)
        try:
            wb = xlrd.open_workbook(filepath)
            parts = []
            for sheet in wb.sheets():
                parts.append(f'[Sheet: {sheet.name}]')
                for row_idx in range(sheet.nrows):
                    row_text = ' | '.join(
                        str(sheet.cell_value(row_idx, c))
                        for c in range(sheet.ncols)
                        if sheet.cell_value(row_idx, c) != ''
                    )
                    if row_text.strip():
                        parts.append(row_text)
            return '\n'.join(parts)
        except Exception:
            pass

        # Fallback: openpyxl (file may be xlsx with wrong extension)
        try:
            wb = load_workbook(filepath, read_only=True, data_only=True)
            parts = []
            for sheet in wb.worksheets:
                parts.append(f'[Sheet: {sheet.title}]')
                for row in sheet.iter_rows(values_only=True):
                    row_text = ' | '.join(str(c) for c in row if c is not None)
                    if row_text.strip():
                        parts.append(row_text)
            return '\n'.join(parts)
        except Exception:
            pass

        # Fallback: HTML (some Excel exports are HTML with .xls extension)
        try:
            raw = open(filepath, 'rb').read()
            enc = chardet.detect(raw)['encoding'] or 'utf-8'
            soup = BeautifulSoup(raw.decode(enc, errors='replace'), 'html.parser')
            text = soup.get_text(separator='\n')
            if text.strip():
                return text
        except Exception:
            pass

        raise ValueError('Could not parse .xls file — try converting it to .xlsx and re-uploading.')

    if ext == 'csv':
        raw = open(filepath, 'rb').read()
        enc = chardet.detect(raw)['encoding'] or 'utf-8'
        text = raw.decode(enc, errors='replace')
        reader = csv.reader(io.StringIO(text))
        return '\n'.join(' | '.join(row) for row in reader if any(row))

    if ext in ('html', 'htm'):
        raw = open(filepath, 'rb').read()
        enc = chardet.detect(raw)['encoding'] or 'utf-8'
        soup = BeautifulSoup(raw.decode(enc, errors='replace'), 'html.parser')
        return soup.get_text(separator='\n')

    if ext == 'json':
        raw = open(filepath, 'rb').read()
        enc = chardet.detect(raw)['encoding'] or 'utf-8'
        data = json.loads(raw.decode(enc, errors='replace'))
        return json.dumps(data, indent=2)

    raw = open(filepath, 'rb').read()
    enc = chardet.detect(raw)['encoding'] or 'utf-8'
    return raw.decode(enc, errors='replace')


def build_prompt(text_content, filename, feature_context=''):
    doc_snippet = text_content[:5000]

    feature_section = ''
    if feature_context and feature_context.strip():
        feature_section = f"""
FEATURE CONTEXT (provided by the tester — use this to sharpen your analysis):
---
{feature_context.strip()}
---
"""

    return f"""You are a senior QA engineer specializing in Amazon Kindle Storefront.
File: "{filename}"
{feature_section}
DOCUMENT (test cases / specs):
---
{doc_snippet}
---

Return ONLY a JSON object (no markdown fences, no commentary) with this exact structure:

{{
  "document_summary": "2-3 sentence description of the document and feature being tested",
  "test_cases_found": [
    {{"id": "TC-001", "title": "title", "category": "category", "description": "what it tests", "priority": "High"}}
  ],
  "coverage_gaps": ["gap 1", "gap 2"],
  "exploratory_scenarios": [
    {{
      "id": "ES-001",
      "title": "scenario title",
      "persona": "Casual Reader",
      "device": "Kindle Paperwhite",
      "scenario": "1. Step one\\n2. Step two\\n3. Step three",
      "what_to_look_for": "Specific observations and potential bugs",
      "risk_level": "High",
      "tags": ["tag1", "tag2"]
    }}
  ],
  "kindle_specific_risks": [
    {{"area": "risk area", "description": "description", "suggested_test": "how to test"}}
  ],
  "total_test_cases": 0,
  "total_scenarios": 0
}}

Rules:
- Extract ALL test cases visible in the document
- If feature context was provided, heavily weight scenarios and gaps toward that specific feature's edge cases, business rules, and integration points
- Generate exactly 8 exploratory scenarios covering different personas (Casual Reader, Power User, Parent, Student, International User, Accessibility User) and devices (Kindle Paperwhite, Kindle App iOS, Kindle App Android, Fire Tablet, All Devices)
- Focus on real Kindle Storefront user journeys: search, browse, purchase, samples, KU, recommendations, wishlists, gifting
- Identify at least 4 coverage gaps (especially for untested edge cases in the described feature) and 3 Kindle-specific risks"""


def call_claude(prompt):
    """Call Claude API. Supports direct Anthropic API key or AWS Bedrock."""
    api_key = os.environ.get('ANTHROPIC_API_KEY', '')
    use_bedrock = os.environ.get('CLAUDE_CODE_USE_BEDROCK', '') == '1' or os.environ.get('USE_BEDROCK', '') == '1'

    if use_bedrock:
        client = anthropic.AnthropicBedrock(
            aws_access_key=os.environ.get('AWS_ACCESS_KEY_ID'),
            aws_secret_key=os.environ.get('AWS_SECRET_ACCESS_KEY'),
            aws_session_token=os.environ.get('AWS_SESSION_TOKEN'),
            aws_region=os.environ.get('AWS_DEFAULT_REGION', 'us-east-1'),
        )
        model = os.environ.get('BEDROCK_MODEL', 'us.anthropic.claude-sonnet-4-5-20251001-v1:0')
    else:
        client = anthropic.Anthropic(api_key=api_key)
        model = os.environ.get('ANTHROPIC_MODEL', 'claude-sonnet-4-6')

    message = client.messages.create(
        model=model,
        max_tokens=4096,
        messages=[{'role': 'user', 'content': prompt}]
    )
    return message.content[0].text.strip()


def run_analysis(job_id, text_content, filename, feature_context=''):
    with _jobs_lock:
        _jobs[job_id]['status'] = 'running'

    try:
        prompt = build_prompt(text_content, filename, feature_context)
        response_text = call_claude(prompt)

        # Extract JSON — handles preamble text + optional markdown fences
        json_match = re.search(r'```(?:json)?\s*(\{[\s\S]+?\})\s*```', response_text)
        if json_match:
            response_text = json_match.group(1)
        else:
            start = response_text.find('{')
            end = response_text.rfind('}')
            if start != -1 and end != -1:
                response_text = response_text[start:end + 1]

        data = json.loads(response_text)
        data['total_test_cases'] = len(data.get('test_cases_found', []))
        data['total_scenarios'] = len(data.get('exploratory_scenarios', []))
        data['filename'] = filename

        with _jobs_lock:
            _jobs[job_id]['status'] = 'done'
            _jobs[job_id]['result'] = data

    except Exception as e:
        with _jobs_lock:
            _jobs[job_id]['status'] = 'error'
            _jobs[job_id]['error'] = str(e)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/analyze', methods=['POST'])
def analyze():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({'error': 'No file selected'}), 400

    if not allowed_file(file.filename):
        return jsonify({
            'error': f'File type not supported. Allowed: {", ".join(sorted(ALLOWED_EXTENSIONS))}'
        }), 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(filepath)

    try:
        text = extract_text_from_file(filepath, filename)
        if not text or not text.strip():
            return jsonify({'error': 'Could not extract text from the file. Is it empty or image-only?'}), 400
    except Exception as e:
        return jsonify({'error': f'File parsing error: {str(e)}'}), 400
    finally:
        try:
            os.remove(filepath)
        except OSError:
            pass

    feature_context = request.form.get('feature_context', '').strip()

    job_id = str(uuid.uuid4())
    with _jobs_lock:
        _jobs[job_id] = {
            'status': 'queued',
            'result': None,
            'error': None,
            'created_at': time.time()
        }

    t = threading.Thread(target=run_analysis, args=(job_id, text, filename, feature_context), daemon=True)
    t.start()

    return jsonify({'job_id': job_id})


@app.route('/result/<job_id>')
def get_result(job_id):
    with _jobs_lock:
        job = _jobs.get(job_id)

    if not job:
        return jsonify({'error': 'Job not found'}), 404

    if job['status'] in ('queued', 'running'):
        return jsonify({'status': job['status']})

    if job['status'] == 'error':
        return jsonify({'status': 'error', 'error': job['error']}), 500

    return jsonify({'status': 'done', 'data': job['result']})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
