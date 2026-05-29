# KindleQA — Storefront Test Analyzer

A web app that reads your test case documents and generates exploratory testing scenarios tailored to real Amazon Kindle Storefront users — powered by Claude AI.

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![Flask](https://img.shields.io/badge/flask-3.x-green)
![Claude](https://img.shields.io/badge/AI-Claude%20Sonnet-orange)

---

## Features

- **Upload any file** — DOCX, PDF, XLSX, XLS, CSV, HTML, TXT, Markdown, JSON, XML, RTF (Quip exports work great)
- **Feature Context box** — describe the feature under test so Claude generates more targeted scenarios and deeper gap analysis
- **Exploratory Scenarios** — 8+ real user journeys across personas (Casual Reader, Power User, Parent, Student, International User, Accessibility User) and devices (Kindle Paperwhite, Fire Tablet, Kindle App iOS/Android/PC)
- **Test Case extraction** — pulls structured test cases from your document with priority tags
- **Coverage Gaps** — highlights what your test plan is missing
- **Kindle-specific Risks** — surfaces risk areas unique to the Kindle ecosystem

---

## Quick Start

### 1. Clone the repo

```bash
git clone https://github.com/YOUR_USERNAME/kindle-test-analyzer.git
cd kindle-test-analyzer
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Set your API key

```bash
cp .env.example .env
# Edit .env and add your ANTHROPIC_API_KEY
```

Or export directly:

```bash
export ANTHROPIC_API_KEY=sk-ant-...
```

### 4. Run

```bash
python app.py
```

Open [http://localhost:8080](http://localhost:8080) in your browser.

---

## Environment Variables

| Variable | Required | Description |
|---|---|---|
| `ANTHROPIC_API_KEY` | Yes (direct API) | Your Anthropic API key from [console.anthropic.com](https://console.anthropic.com) |
| `ANTHROPIC_MODEL` | No | Model to use (default: `claude-sonnet-4-6`) |
| `USE_BEDROCK` | No | Set to `1` to use AWS Bedrock instead of direct API |
| `BEDROCK_MODEL` | No | Bedrock model ID (default: `us.anthropic.claude-sonnet-4-5-20251001-v1:0`) |
| `AWS_ACCESS_KEY_ID` | Bedrock only | AWS credentials |
| `AWS_SECRET_ACCESS_KEY` | Bedrock only | AWS credentials |
| `AWS_SESSION_TOKEN` | Bedrock only | AWS session token (if using temporary credentials) |
| `AWS_DEFAULT_REGION` | Bedrock only | AWS region (default: `us-east-1`) |

---

## Project Structure

```
kindle-test-analyzer/
├── app.py                  # Flask backend
├── templates/
│   └── index.html          # Single-page frontend
├── requirements.txt
├── .env.example
├── .gitignore
└── README.md
```

---

## How It Works

1. Upload a test case document (Quip export, Word doc, spreadsheet, etc.)
2. Optionally describe the feature being tested in the Feature Context box
3. The backend extracts all text from the file, builds a prompt, and sends it to Claude
4. Claude analyzes the test cases, finds coverage gaps, and generates exploratory scenarios
5. Results are polled asynchronously and rendered in a tabbed dashboard

---

## Supported File Formats

| Format | Notes |
|---|---|
| `.docx` / `.doc` | Full text + table extraction |
| `.pdf` | Text layer extraction (not image-based PDFs) |
| `.xlsx` | All sheets |
| `.xls` | Legacy Excel; falls back to openpyxl or HTML parser if needed |
| `.csv` | Auto-detects encoding |
| `.html` / `.htm` | Quip exports, wiki pages |
| `.txt` / `.md` | Plain text |
| `.json` | Pretty-printed for readability |
| `.xml` / `.rtf` | Plain text extraction |

---

## License

MIT
