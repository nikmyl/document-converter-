#!/usr/bin/env python3
"""
Flask Web Application for Document Converter
Converts between Markdown (.md), Word documents (.docx), and PDF files
Provides a web interface with drag-and-drop functionality
Supports batch conversion with multiple files or zip upload
"""

import os
import io
import tempfile
import zipfile
import shutil
from pathlib import Path
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from md_to_docx import MarkdownToDocxConverter
from docx_to_md import DocxToMarkdownConverter

# PDF converters - imported lazily to handle missing GTK+ on Windows
MarkdownToPdfConverter = None
PdfToMarkdownConverter = None
DocxToPdfConverter = None
PdfToDocxConverter = None

def _load_pdf_converters():
    """Load PDF converters lazily"""
    global MarkdownToPdfConverter, PdfToMarkdownConverter, DocxToPdfConverter, PdfToDocxConverter
    if MarkdownToPdfConverter is None:
        try:
            from md_to_pdf import MarkdownToPdfConverter as _MdToPdf
            from pdf_to_md import PdfToMarkdownConverter as _PdfToMd
            from docx_to_pdf import DocxToPdfConverter as _DocxToPdf
            from pdf_to_docx import PdfToDocxConverter as _PdfToDocx
            MarkdownToPdfConverter = _MdToPdf
            PdfToMarkdownConverter = _PdfToMd
            DocxToPdfConverter = _DocxToPdf
            PdfToDocxConverter = _PdfToDocx
            return True
        except Exception as e:
            print(f"Warning: PDF converters not available: {e}")
            print("To enable PDF support on Windows, install GTK+ runtime:")
            print("  https://github.com/nickvidal/msys2/wiki/MSYS2-installation")
            return False
    return True

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max for batch uploads
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()

# Supported file extensions
MARKDOWN_EXTENSIONS = {'md', 'markdown', 'txt'}
DOCX_EXTENSIONS = {'docx'}
PDF_EXTENSIONS = {'pdf'}
ZIP_EXTENSIONS = {'zip'}
ALLOWED_EXTENSIONS = MARKDOWN_EXTENSIONS | DOCX_EXTENSIONS | PDF_EXTENSIONS | ZIP_EXTENSIONS


def allowed_file(filename):
    """Check if the uploaded file has an allowed extension"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def is_convertible_file(filename):
    """Check if the file can be converted (not a zip)"""
    ext = get_file_extension(filename)
    return ext in MARKDOWN_EXTENSIONS or ext in DOCX_EXTENSIONS or ext in PDF_EXTENSIONS


def get_file_extension(filename):
    """Get the lowercase file extension"""
    return filename.rsplit('.', 1)[1].lower() if '.' in filename else ''


def is_markdown_file(filename):
    """Check if the file is a markdown file"""
    return get_file_extension(filename) in MARKDOWN_EXTENSIONS


def is_docx_file(filename):
    """Check if the file is a Word document"""
    return get_file_extension(filename) in DOCX_EXTENSIONS


def is_pdf_file(filename):
    """Check if the file is a PDF document"""
    return get_file_extension(filename) in PDF_EXTENSIONS


def is_zip_file(filename):
    """Check if the file is a zip archive"""
    return get_file_extension(filename) in ZIP_EXTENSIONS


def get_converter_and_output(input_path, filename, output_dir, target_format=None):
    """
    Get the appropriate converter and output path for a file

    Args:
        input_path: Path to input file
        filename: Original filename
        output_dir: Base output directory
        target_format: Target format ('pdf', 'docx', 'md') or None for default

    Returns:
        Tuple of (converter, output_path, direction) or raises ValueError
    """
    if is_markdown_file(filename):
        if target_format == 'pdf':
            if not _load_pdf_converters():
                raise ValueError('PDF conversion not available. GTK+ libraries required on Windows.')
            output_subdir = os.path.join(output_dir, 'PDF')
            os.makedirs(output_subdir, exist_ok=True)
            output_filename = Path(filename).stem + '.pdf'
            output_path = os.path.join(output_subdir, output_filename)
            return MarkdownToPdfConverter(input_path, output_path), output_path, 'md_to_pdf'
        else:
            output_subdir = os.path.join(output_dir, 'DOCX')
            os.makedirs(output_subdir, exist_ok=True)
            output_filename = Path(filename).stem + '.docx'
            output_path = os.path.join(output_subdir, output_filename)
            return MarkdownToDocxConverter(input_path, output_path), output_path, 'md_to_docx'

    elif is_docx_file(filename):
        if target_format == 'pdf':
            if not _load_pdf_converters():
                raise ValueError('PDF conversion not available. GTK+ libraries required on Windows.')
            output_subdir = os.path.join(output_dir, 'PDF')
            os.makedirs(output_subdir, exist_ok=True)
            output_filename = Path(filename).stem + '.pdf'
            output_path = os.path.join(output_subdir, output_filename)
            return DocxToPdfConverter(input_path, output_path), output_path, 'docx_to_pdf'
        else:
            output_subdir = os.path.join(output_dir, 'MD')
            os.makedirs(output_subdir, exist_ok=True)
            output_filename = Path(filename).stem + '.md'
            output_path = os.path.join(output_subdir, output_filename)
            return DocxToMarkdownConverter(input_path, output_path), output_path, 'docx_to_md'

    elif is_pdf_file(filename):
        if not _load_pdf_converters():
            raise ValueError('PDF conversion not available. GTK+ libraries required on Windows.')
        if target_format == 'docx':
            output_subdir = os.path.join(output_dir, 'DOCX')
            os.makedirs(output_subdir, exist_ok=True)
            output_filename = Path(filename).stem + '.docx'
            output_path = os.path.join(output_subdir, output_filename)
            return PdfToDocxConverter(input_path, output_path), output_path, 'pdf_to_docx'
        else:
            output_subdir = os.path.join(output_dir, 'MD')
            os.makedirs(output_subdir, exist_ok=True)
            output_filename = Path(filename).stem + '.md'
            output_path = os.path.join(output_subdir, output_filename)
            return PdfToMarkdownConverter(input_path, output_path), output_path, 'pdf_to_md'

    else:
        raise ValueError(f'Unsupported file type: {filename}')


def get_mimetype_for_direction(direction):
    """Get MIME type for output based on conversion direction"""
    mimetypes = {
        'md_to_docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'md_to_pdf': 'application/pdf',
        'docx_to_md': 'text/markdown',
        'docx_to_pdf': 'application/pdf',
        'pdf_to_md': 'text/markdown',
        'pdf_to_docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    }
    return mimetypes.get(direction, 'application/octet-stream')


def convert_single_file(input_path, output_dir, target_format=None):
    """
    Convert a single file and return the output path and metadata

    Args:
        input_path: Path to input file
        output_dir: Directory to save output (with MD/, DOCX/, PDF/ subdirs)
        target_format: Target format ('pdf', 'docx', 'md') or None for default

    Returns:
        Tuple of (output_path, direction) or (None, error_message)
    """
    filename = os.path.basename(input_path)

    try:
        converter, output_path, direction = get_converter_and_output(
            input_path, filename, output_dir, target_format
        )
        converter.convert()
        return output_path, direction
    except ValueError as e:
        return None, str(e)
    except Exception as e:
        return None, str(e)


def create_output_zip(output_dir, include_root=False):
    """
    Create a zip file from the output directory

    Args:
        output_dir: Directory containing MD/ and DOCX/ folders
        include_root: Whether to include root folder in zip paths

    Returns:
        BytesIO object containing the zip file
    """
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(output_dir):
            for file in files:
                file_path = os.path.join(root, file)
                # Calculate archive path (relative to output_dir)
                arcname = os.path.relpath(file_path, output_dir)
                zf.write(file_path, arcname)

    zip_buffer.seek(0)
    return zip_buffer


@app.route('/')
def index():
    """Render the main page"""
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert_file():
    """
    Handle single file upload and conversion

    Query Parameters:
        format: Target format ('pdf', 'docx', 'md') - optional
    """
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['file']

        # Check if filename is empty
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        # Validate file extension
        if not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type. Please upload a .md, .markdown, .txt, .docx, or .pdf file'}), 400

        # Get target format from query parameter
        target_format = request.args.get('format', None)

        # Secure the filename
        filename = secure_filename(file.filename)

        # Create temporary directory
        temp_dir = tempfile.mkdtemp()
        input_path = os.path.join(temp_dir, filename)

        # Save uploaded file
        file.save(input_path)

        # Get converter and output info
        try:
            converter, output_path, direction = get_converter_and_output(
                input_path, filename, temp_dir, target_format
            )
        except ValueError as e:
            return jsonify({'error': str(e)}), 400

        # Convert
        converter.convert()

        # Get output filename and mimetype
        output_filename = os.path.basename(output_path)
        mimetype = get_mimetype_for_direction(direction)

        # Send the converted file
        return send_file(
            output_path,
            as_attachment=True,
            download_name=output_filename,
            mimetype=mimetype
        )

    except Exception as e:
        return jsonify({'error': f'Conversion failed: {str(e)}'}), 500


@app.route('/convert-batch', methods=['POST'])
def convert_batch():
    """
    Handle batch file conversion (multiple files or zip)

    Accepts:
        - Multiple files via 'files' field
        - Single zip file via 'file' field

    Returns:
        - Zip file containing converted files organized in MD/ and DOCX/ folders
    """
    try:
        temp_dir = tempfile.mkdtemp()
        input_dir = os.path.join(temp_dir, 'input')
        output_dir = os.path.join(temp_dir, 'output')
        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        files_to_convert = []
        results = {'converted': [], 'errors': [], 'skipped': []}

        # Check for multiple files
        if 'files' in request.files:
            files = request.files.getlist('files')
            for file in files:
                if file.filename == '':
                    continue
                filename = secure_filename(file.filename)
                if is_convertible_file(filename):
                    file_path = os.path.join(input_dir, filename)
                    file.save(file_path)
                    files_to_convert.append(file_path)
                else:
                    results['skipped'].append(filename)

        # Check for single zip file
        elif 'file' in request.files:
            file = request.files['file']
            if file.filename == '':
                return jsonify({'error': 'No file selected'}), 400

            filename = secure_filename(file.filename)

            if is_zip_file(filename):
                # Extract zip file
                zip_path = os.path.join(temp_dir, filename)
                file.save(zip_path)

                with zipfile.ZipFile(zip_path, 'r') as zf:
                    for zip_info in zf.infolist():
                        if zip_info.is_dir():
                            continue
                        # Get just the filename (ignore folder structure)
                        extracted_name = os.path.basename(zip_info.filename)
                        if not extracted_name:
                            continue
                        extracted_name = secure_filename(extracted_name)
                        if is_convertible_file(extracted_name):
                            # Extract to input dir
                            extracted_path = os.path.join(input_dir, extracted_name)
                            with zf.open(zip_info) as src, open(extracted_path, 'wb') as dst:
                                dst.write(src.read())
                            files_to_convert.append(extracted_path)
                        elif extracted_name:
                            results['skipped'].append(extracted_name)
            elif is_convertible_file(filename):
                # Single convertible file - redirect to single convert
                file_path = os.path.join(input_dir, filename)
                file.save(file_path)
                files_to_convert.append(file_path)
            else:
                return jsonify({'error': f'Unsupported file type: {filename}'}), 400
        else:
            return jsonify({'error': 'No files uploaded'}), 400

        if not files_to_convert:
            return jsonify({'error': 'No convertible files found'}), 400

        # Convert all files
        for file_path in files_to_convert:
            output_path, result = convert_single_file(file_path, output_dir)
            filename = os.path.basename(file_path)

            if output_path:
                results['converted'].append({
                    'input': filename,
                    'output': os.path.basename(output_path),
                    'direction': result
                })
            else:
                results['errors'].append({
                    'file': filename,
                    'error': result
                })

        # If only one file was converted and no errors, return the single file
        if len(results['converted']) == 1 and not results['errors']:
            converted = results['converted'][0]
            direction = converted['direction']

            # Determine output folder based on direction
            if direction in ('md_to_docx', 'pdf_to_docx'):
                output_folder = 'DOCX'
            elif direction in ('md_to_pdf', 'docx_to_pdf'):
                output_folder = 'PDF'
            else:
                output_folder = 'MD'

            output_path = os.path.join(output_dir, output_folder, converted['output'])
            mimetype = get_mimetype_for_direction(direction)

            return send_file(
                output_path,
                as_attachment=True,
                download_name=converted['output'],
                mimetype=mimetype
            )

        # Create and send zip file
        zip_buffer = create_output_zip(output_dir)

        # Clean up temp directory (will be handled by OS, but good practice)
        # Note: In production, you'd want to use a proper cleanup mechanism

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name='converted_files.zip',
            mimetype='application/zip'
        )

    except Exception as e:
        return jsonify({'error': f'Batch conversion failed: {str(e)}'}), 500


@app.route('/health')
def health():
    """Health check endpoint"""
    return jsonify({'status': 'ok'})


@app.errorhandler(413)
def too_large(e):
    """Handle file too large error"""
    return jsonify({'error': 'File too large. Maximum size is 100MB for batch uploads'}), 413


if __name__ == '__main__':
    print("=" * 60)
    print("Document Converter - Web Interface")
    print("Supports: Markdown <-> Word (DOCX) <-> PDF")
    print("=" * 60)
    print("\nStarting server...")
    print("Access the application at: http://localhost:5000")
    print("\nSupported conversions:")
    print("  .md, .markdown, .txt  ->  .docx or .pdf")
    print("  .docx                 ->  .md or .pdf")
    print("  .pdf                  ->  .md or .docx")
    print("\nBatch conversion:")
    print("  - Upload multiple files at once")
    print("  - Upload a .zip file containing documents")
    print("  - Output organized in MD/, DOCX/, and PDF/ folders")
    print("\nPress CTRL+C to stop the server")
    print("=" * 60)

    app.run(debug=True, host='0.0.0.0', port=5000)
