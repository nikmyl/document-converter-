# CLAUDE.md

This file provides context for AI assistants working with this codebase.

## Project Overview

A document converter supporting bidirectional conversion between Markdown (.md), Microsoft Word (.docx), and PDF formats. All 6 conversion directions are supported:
- MD ↔ DOCX
- MD ↔ PDF
- DOCX ↔ PDF

Provides two interfaces:
- Command-line interface (CLI)
- Flask web application with drag-and-drop

## Project Structure

```
MD to Word/
├── converter.py          # Unified CLI (auto-detects direction, supports folders)
├── md_to_docx.py         # Markdown → DOCX converter class
├── docx_to_md.py         # DOCX → Markdown converter class
├── md_to_pdf.py          # Markdown → PDF converter class (ReportLab)
├── pdf_to_md.py          # PDF → Markdown converter class (pdfplumber)
├── docx_to_pdf.py        # DOCX → PDF converter (chains DOCX→MD→PDF)
├── pdf_to_docx.py        # PDF → DOCX converter (chains PDF→MD→DOCX)
├── app.py                # Flask web application (single + batch)
├── requirements.txt      # Python dependencies
├── templates/
│   └── index.html        # Web UI template (Jinja2)
├── static/
│   ├── css/style.css     # Web UI styling
│   └── js/script.js      # Client-side JavaScript (handles batch mode)
├── start_webapp.bat      # Windows launcher for web app
├── start_webapp.sh       # Linux/Mac launcher for web app
├── README.md             # Project readme
├── DOCUMENTATION.md      # Detailed documentation
├── QUICKSTART.md         # Quick start guide
├── WEB_APP_GUIDE.md      # Web application usage guide
├── sample.md / sample.docx      # Test files
└── table_test.md / table_test.docx  # Table conversion test files
```

**Generated output folders** (created by converter):
- `MD/` - Markdown files converted from DOCX or PDF
- `DOCX/` - Word files converted from Markdown or PDF
- `PDF/` - PDF files converted from Markdown or DOCX

## Key Components

### Converters

- **`MarkdownToDocxConverter`** (md_to_docx.py): Parses markdown line-by-line, handles headings, bold/italic, lists, code blocks, tables, blockquotes, links. Accepts `.md`, `.markdown`, `.txt` files.
- **`DocxToMarkdownConverter`** (docx_to_md.py): Reads DOCX via python-docx, extracts styles and formatting, outputs markdown.
- **`MarkdownToPdfConverter`** (md_to_pdf.py): Converts markdown to PDF using ReportLab. Supports headings, tables, lists, code blocks, inline formatting.
- **`PdfToMarkdownConverter`** (pdf_to_md.py): Extracts text and tables from PDF using pdfplumber/pypdf. Detects headings via font size analysis.
- **`DocxToPdfConverter`** (docx_to_pdf.py): Chain converter (DOCX → MD → PDF).
- **`PdfToDocxConverter`** (pdf_to_docx.py): Chain converter (PDF → MD → DOCX).

### Web Application

- Flask app on port 5000
- `/convert` POST endpoint - single file conversion with `?format=` parameter
- `/convert-batch` POST endpoint - multiple files or ZIP upload
- Auto-detects source format, shows two target format buttons
- Returns converted file(s) as download (ZIP for batch)

### CLI Features

- Auto-detects conversion direction from file extension
- **Target format flags**: `--to-pdf`, `--to-docx`, `--to-md`
- **Folder support**: Converts all files in a directory
- **Recursive mode**: `-r` flag for subdirectories
- **Output organization**: Creates `MD/`, `DOCX/`, and `PDF/` subfolders

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Single file conversion (CLI) - default targets
python converter.py document.md      # → document.docx (default)
python converter.py document.docx    # → document.md (default)
python converter.py document.pdf     # → document.md (default)

# Explicit target format
python converter.py document.md --to-pdf     # → document.pdf
python converter.py document.docx --to-pdf   # → document.pdf
python converter.py document.pdf --to-docx   # → document.docx

# Folder conversion (creates MD/, DOCX/, PDF/ subfolders)
python converter.py ./my_folder      # Convert all files in folder
python converter.py ./my_folder -r   # Recursive conversion

# Batch convert with wildcards
python converter.py *.md             # All markdown files
python converter.py *.pdf --to-docx  # All PDFs to DOCX

# Start web app
python app.py  # http://localhost:5000
```

## Folder Conversion Output

```
Input Folder/
├── doc1.md
├── doc2.docx
├── report.pdf

After running: python converter.py ./Input_Folder

Input Folder/
├── MD/                  # Converted from DOCX/PDF
│   ├── doc2.md
│   └── report.md
├── DOCX/                # Converted from MD
│   └── doc1.docx
├── PDF/                 # (if --to-pdf used)
├── doc1.md              # Original files unchanged
├── doc2.docx
└── report.pdf
```

## Web Interface Modes

1. **Single File Mode** (default)
   - Drag-and-drop or browse for one file
   - Shows **two convert buttons** based on source type:
     - MD file → "Convert to DOCX" + "Convert to PDF"
     - DOCX file → "Convert to MD" + "Convert to PDF"
     - PDF file → "Convert to MD" + "Convert to DOCX"
   - Downloads converted file directly

2. **Batch / Folder Mode**
   - Toggle via button at top
   - Accept multiple files or ZIP archive
   - Shows file list with conversion counts (MD, DOCX, PDF)
   - Downloads results as organized ZIP (MD/, DOCX/, PDF/ folders)

## Dependencies

- `python-docx>=0.8.11` - Word document creation/reading
- `markdown>=3.4.0` - Markdown processing
- `Flask>=2.3.0` - Web framework
- `reportlab>=4.0.0` - PDF creation (pure Python, no external deps)
- `pdfplumber>=0.10.0` - PDF table extraction
- `pypdf>=4.0.0` - PDF text extraction

## File Extension Mapping

| Input Extensions | Default Output | Converter Class |
|-----------------|----------------|-----------------|
| .md, .markdown, .txt | .docx | MarkdownToDocxConverter |
| .md, .markdown, .txt | .pdf (with --to-pdf) | MarkdownToPdfConverter |
| .docx | .md | DocxToMarkdownConverter |
| .docx | .pdf (with --to-pdf) | DocxToPdfConverter |
| .pdf | .md | PdfToMarkdownConverter |
| .pdf | .docx (with --to-docx) | PdfToDocxConverter |

**Note**: `.txt` files are excluded from automatic folder scanning (too generic) but can be converted explicitly.

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/` | GET | Web interface |
| `/convert?format=<target>` | POST | Single file conversion (field: `file`, format: `md`, `docx`, `pdf`) |
| `/convert-batch` | POST | Batch conversion (field: `files` or `file` for ZIP) |
| `/health` | GET | Health check |

## Testing

Test files included:
- `sample.md` / `sample.docx` - Basic formatting test
- `table_test.md` / `table_test.docx` - Table conversion test

```bash
# Test single file
python converter.py sample.md

# Test PDF conversion
python converter.py sample.md --to-pdf

# Test folder conversion
python converter.py . -v

# Test round-trip
python converter.py sample.md -o test.docx
python converter.py test.docx -o test_roundtrip.md
```

## Known Limitations

- Images are not embedded (link text only)
- No syntax highlighting for code blocks
- HTML within markdown is not processed
- Complex nested formatting may not render perfectly
- Hyperlink extraction from DOCX depends on relationship IDs
- **PDF extraction**: Scanned PDFs (images) not supported (no OCR)
- **PDF extraction**: Complex multi-column layouts may not extract perfectly
- **PDF creation**: ReportLab doesn't support strikethrough in paragraphs

## Development Notes

- Web UI uses vanilla JavaScript (no frameworks)
- Mode toggle switches between single/batch file handling
- Drag-and-drop supports multiple files (auto-switches to batch mode)
- Temporary files created in system temp directory during web conversions
- ZIP output preserves MD/, DOCX/, and PDF/ folder structure
- Max file size: 16MB single, 100MB batch
- PDF converters use lazy loading in app.py for graceful degradation
- ReportLab chosen over WeasyPrint to avoid GTK+ dependency on Windows
