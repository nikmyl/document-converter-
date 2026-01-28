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
import atexit
import threading
import time
from pathlib import Path
from functools import wraps
from flask import Flask, render_template, request, send_file, jsonify
from werkzeug.utils import secure_filename
from md_to_docx import MarkdownToDocxConverter
from docx_to_md import DocxToMarkdownConverter

# File magic bytes for content validation
FILE_SIGNATURES = {
    'docx': [b'PK\x03\x04'],  # ZIP-based format (DOCX, XLSX, etc.)
    'pdf': [b'%PDF'],
    'zip': [b'PK\x03\x04', b'PK\x05\x06'],  # Standard ZIP and empty ZIP
}

# Rate limiting storage
_rate_limit_store = {}
_rate_limit_lock = threading.Lock()

# Temp directories to clean up
_temp_dirs_to_cleanup = set()
_cleanup_lock = threading.Lock()

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
app.config['MAX_SINGLE_FILE_SIZE'] = 16 * 1024 * 1024  # 16MB max for single files
app.config['RATE_LIMIT_REQUESTS'] = 30  # Max requests per window
app.config['RATE_LIMIT_WINDOW'] = 60  # Window in seconds
app.config['MAX_ZIP_RATIO'] = 100  # Max decompression ratio (ZIP bomb protection)

# Supported file extensions
MARKDOWN_EXTENSIONS = {'md', 'markdown', 'txt'}
DOCX_EXTENSIONS = {'docx'}
PDF_EXTENSIONS = {'pdf'}
ZIP_EXTENSIONS = {'zip'}
ALLOWED_EXTENSIONS = MARKDOWN_EXTENSIONS | DOCX_EXTENSIONS | PDF_EXTENSIONS | ZIP_EXTENSIONS


def cleanup_temp_dirs():
    """Clean up temporary directories on exit"""
    with _cleanup_lock:
        for temp_dir in list(_temp_dirs_to_cleanup):
            try:
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception:
                pass
        _temp_dirs_to_cleanup.clear()


# Register cleanup function
atexit.register(cleanup_temp_dirs)


def create_temp_dir():
    """Create a temp directory and track it for cleanup"""
    temp_dir = tempfile.mkdtemp()
    with _cleanup_lock:
        _temp_dirs_to_cleanup.add(temp_dir)
    return temp_dir


def cleanup_single_temp_dir(temp_dir):
    """Clean up a single temp directory after use"""
    try:
        with _cleanup_lock:
            _temp_dirs_to_cleanup.discard(temp_dir)
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)
    except Exception:
        pass


def rate_limit(f):
    """Rate limiting decorator"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        client_ip = request.remote_addr or 'unknown'
        current_time = time.time()
        window = app.config['RATE_LIMIT_WINDOW']
        max_requests = app.config['RATE_LIMIT_REQUESTS']

        with _rate_limit_lock:
            # Clean up old entries
            if client_ip in _rate_limit_store:
                _rate_limit_store[client_ip] = [
                    t for t in _rate_limit_store[client_ip]
                    if current_time - t < window
                ]
            else:
                _rate_limit_store[client_ip] = []

            # Check rate limit
            if len(_rate_limit_store[client_ip]) >= max_requests:
                return jsonify({'error': 'Too many requests. Please try again later.'}), 429

            # Record this request
            _rate_limit_store[client_ip].append(current_time)

        return f(*args, **kwargs)
    return decorated_function


def validate_file_content(file_obj, expected_type):
    """
    Validate file content by checking magic bytes

    Args:
        file_obj: File object to validate
        expected_type: Expected file type ('docx', 'pdf', 'zip')

    Returns:
        bool: True if valid, False otherwise
    """
    if expected_type not in FILE_SIGNATURES:
        # For markdown/text files, we don't check magic bytes
        return True

    # Read first few bytes
    file_obj.seek(0)
    header = file_obj.read(8)
    file_obj.seek(0)

    signatures = FILE_SIGNATURES[expected_type]
    for sig in signatures:
        if header.startswith(sig):
            return True

    return False


def is_safe_zip_entry(zip_entry_name, extract_to):
    """
    Check if a ZIP entry is safe to extract (no path traversal)

    Args:
        zip_entry_name: Name of the entry in the ZIP file
        extract_to: Base directory for extraction

    Returns:
        bool: True if safe, False if path traversal detected
    """
    # Normalize the path
    extract_to = os.path.abspath(extract_to)

    # Get the target path
    target_path = os.path.abspath(os.path.join(extract_to, zip_entry_name))

    # Check if target is within extract directory
    return target_path.startswith(extract_to + os.sep) or target_path == extract_to


def check_zip_bomb(zip_file, max_ratio=None):
    """
    Check for ZIP bomb by comparing compressed vs uncompressed size

    Args:
        zip_file: ZipFile object
        max_ratio: Maximum allowed compression ratio

    Returns:
        bool: True if safe, False if potential ZIP bomb
    """
    if max_ratio is None:
        max_ratio = app.config['MAX_ZIP_RATIO']

    total_compressed = 0
    total_uncompressed = 0

    for info in zip_file.infolist():
        total_compressed += info.compress_size
        total_uncompressed += info.file_size

    if total_compressed == 0:
        return True

    ratio = total_uncompressed / total_compressed
    return ratio <= max_ratio


def sanitize_error_message(error):
    """
    Sanitize error messages to avoid leaking internal details

    Args:
        error: Exception or error string

    Returns:
        str: Safe error message
    """
    error_str = str(error).lower()

    # Check for sensitive information patterns
    sensitive_patterns = [
        'traceback', 'file "/', 'line ', 'in <module>',
        '/home/', '/users/', 'c:\\', 'd:\\', 'password',
        'secret', 'key', 'token', 'credential'
    ]

    for pattern in sensitive_patterns:
        if pattern in error_str:
            return 'An error occurred during conversion. Please try again.'

    # Limit error message length
    error_msg = str(error)
    if len(error_msg) > 200:
        error_msg = error_msg[:200] + '...'

    return error_msg


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
@rate_limit
def convert_file():
    """
    Handle single file upload and conversion

    Query Parameters:
        format: Target format ('pdf', 'docx', 'md') - optional
    """
    temp_dir = None
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

        # Check file size (single file limit)
        file.seek(0, 2)  # Seek to end
        file_size = file.tell()
        file.seek(0)  # Seek back to start

        if file_size > app.config['MAX_SINGLE_FILE_SIZE']:
            return jsonify({'error': 'File too large. Maximum size is 16MB for single files.'}), 413

        # Validate file content (magic bytes) for binary files
        ext = get_file_extension(file.filename)
        if ext in ('docx', 'pdf'):
            if not validate_file_content(file, ext):
                return jsonify({'error': 'Invalid file content. The file may be corrupted or not a valid format.'}), 400

        # Get target format from query parameter
        target_format = request.args.get('format', None)

        # Secure the filename
        filename = secure_filename(file.filename)
        if not filename:
            return jsonify({'error': 'Invalid filename'}), 400

        # Create temporary directory (tracked for cleanup)
        temp_dir = create_temp_dir()
        input_path = os.path.join(temp_dir, filename)

        # Save uploaded file
        file.save(input_path)

        # Get converter and output info
        try:
            converter, output_path, direction = get_converter_and_output(
                input_path, filename, temp_dir, target_format
            )
        except ValueError as e:
            return jsonify({'error': sanitize_error_message(e)}), 400

        # Convert
        converter.convert()

        # Get output filename and mimetype
        output_filename = os.path.basename(output_path)
        mimetype = get_mimetype_for_direction(direction)

        # Read file into memory before cleanup
        with open(output_path, 'rb') as f:
            output_data = io.BytesIO(f.read())
        output_data.seek(0)

        # Clean up temp directory
        cleanup_single_temp_dir(temp_dir)
        temp_dir = None

        # Send the converted file
        return send_file(
            output_data,
            as_attachment=True,
            download_name=output_filename,
            mimetype=mimetype
        )

    except Exception as e:
        if temp_dir:
            cleanup_single_temp_dir(temp_dir)
        return jsonify({'error': f'Conversion failed: {sanitize_error_message(e)}'}), 500


@app.route('/convert-batch', methods=['POST'])
@rate_limit
def convert_batch():
    """
    Handle batch file conversion (multiple files or zip)

    Accepts:
        - Multiple files via 'files' field
        - Single zip file via 'file' field

    Returns:
        - Zip file containing converted files organized in MD/ and DOCX/ folders
    """
    temp_dir = None
    try:
        temp_dir = create_temp_dir()
        input_dir = os.path.join(temp_dir, 'input')
        output_dir = os.path.join(temp_dir, 'output')
        os.makedirs(input_dir, exist_ok=True)
        os.makedirs(output_dir, exist_ok=True)

        files_to_convert = []
        results = {'converted': [], 'errors': [], 'skipped': []}
        seen_filenames = set()  # Track filenames to prevent overwrites

        # Check for multiple files
        if 'files' in request.files:
            files = request.files.getlist('files')
            for file in files:
                if file.filename == '':
                    continue

                # Check individual file size
                file.seek(0, 2)
                file_size = file.tell()
                file.seek(0)

                if file_size > app.config['MAX_SINGLE_FILE_SIZE']:
                    results['errors'].append({
                        'file': file.filename,
                        'error': 'File too large (max 16MB per file)'
                    })
                    continue

                filename = secure_filename(file.filename)
                if not filename:
                    continue

                # Handle duplicate filenames
                if filename in seen_filenames:
                    base, ext = os.path.splitext(filename)
                    counter = 1
                    while f"{base}_{counter}{ext}" in seen_filenames:
                        counter += 1
                    filename = f"{base}_{counter}{ext}"

                seen_filenames.add(filename)

                if is_convertible_file(filename):
                    # Validate file content for binary files
                    ext = get_file_extension(filename)
                    if ext in ('docx', 'pdf') and not validate_file_content(file, ext):
                        results['errors'].append({
                            'file': filename,
                            'error': 'Invalid file content'
                        })
                        continue

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
            if not filename:
                return jsonify({'error': 'Invalid filename'}), 400

            if is_zip_file(filename):
                # Validate ZIP file content
                if not validate_file_content(file, 'zip'):
                    return jsonify({'error': 'Invalid ZIP file'}), 400

                # Extract zip file
                zip_path = os.path.join(temp_dir, filename)
                file.save(zip_path)

                try:
                    with zipfile.ZipFile(zip_path, 'r') as zf:
                        # Check for ZIP bomb
                        if not check_zip_bomb(zf):
                            return jsonify({'error': 'ZIP file rejected: suspicious compression ratio'}), 400

                        for zip_info in zf.infolist():
                            if zip_info.is_dir():
                                continue

                            # Check for path traversal
                            if not is_safe_zip_entry(zip_info.filename, input_dir):
                                results['errors'].append({
                                    'file': zip_info.filename,
                                    'error': 'Invalid path in ZIP'
                                })
                                continue

                            # Check individual file size within ZIP
                            if zip_info.file_size > app.config['MAX_SINGLE_FILE_SIZE']:
                                results['errors'].append({
                                    'file': zip_info.filename,
                                    'error': 'File too large (max 16MB per file)'
                                })
                                continue

                            # Get just the filename (ignore folder structure)
                            extracted_name = os.path.basename(zip_info.filename)
                            if not extracted_name:
                                continue
                            extracted_name = secure_filename(extracted_name)
                            if not extracted_name:
                                continue

                            # Handle duplicate filenames
                            if extracted_name in seen_filenames:
                                base, ext = os.path.splitext(extracted_name)
                                counter = 1
                                while f"{base}_{counter}{ext}" in seen_filenames:
                                    counter += 1
                                extracted_name = f"{base}_{counter}{ext}"

                            seen_filenames.add(extracted_name)

                            if is_convertible_file(extracted_name):
                                # Extract to input dir (safely read in chunks)
                                extracted_path = os.path.join(input_dir, extracted_name)
                                with zf.open(zip_info) as src:
                                    with open(extracted_path, 'wb') as dst:
                                        # Read in 64KB chunks to avoid memory issues
                                        chunk_size = 65536
                                        bytes_read = 0
                                        while True:
                                            chunk = src.read(chunk_size)
                                            if not chunk:
                                                break
                                            dst.write(chunk)
                                            bytes_read += len(chunk)
                                            # Safety check
                                            if bytes_read > app.config['MAX_SINGLE_FILE_SIZE']:
                                                break
                                files_to_convert.append(extracted_path)
                            elif extracted_name:
                                results['skipped'].append(extracted_name)
                except zipfile.BadZipFile:
                    return jsonify({'error': 'Invalid or corrupted ZIP file'}), 400
            elif is_convertible_file(filename):
                # Single convertible file - redirect to single convert
                file_path = os.path.join(input_dir, filename)
                file.save(file_path)
                files_to_convert.append(file_path)
            else:
                return jsonify({'error': f'Unsupported file type'}), 400
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
                    'error': sanitize_error_message(result)
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

            # Read file into memory before cleanup
            with open(output_path, 'rb') as f:
                output_data = io.BytesIO(f.read())
            output_data.seek(0)

            # Clean up temp directory
            cleanup_single_temp_dir(temp_dir)
            temp_dir = None

            return send_file(
                output_data,
                as_attachment=True,
                download_name=converted['output'],
                mimetype=mimetype
            )

        # Create and send zip file
        zip_buffer = create_output_zip(output_dir)

        # Clean up temp directory
        cleanup_single_temp_dir(temp_dir)
        temp_dir = None

        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name='converted_files.zip',
            mimetype='application/zip'
        )

    except Exception as e:
        if temp_dir:
            cleanup_single_temp_dir(temp_dir)
        return jsonify({'error': f'Batch conversion failed: {sanitize_error_message(e)}'}), 500


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
