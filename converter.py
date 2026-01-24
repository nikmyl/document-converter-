#!/usr/bin/env python3
"""
Bidirectional Document Converter
Converts between Markdown (.md), Word documents (.docx), and PDF files
Auto-detects conversion direction based on file extension
Supports folder batch conversion with organized output
"""

import sys
import argparse
import os
from pathlib import Path
from typing import List, Tuple, Optional

from md_to_docx import MarkdownToDocxConverter
from docx_to_md import DocxToMarkdownConverter
from md_to_pdf import MarkdownToPdfConverter
from pdf_to_md import PdfToMarkdownConverter
from docx_to_pdf import DocxToPdfConverter
from pdf_to_docx import PdfToDocxConverter

# Supported extensions
MARKDOWN_EXTENSIONS = {'.md', '.markdown', '.txt'}
MARKDOWN_EXTENSIONS_FOLDER = {'.md', '.markdown'}  # Exclude .txt from folder scan (too generic)
DOCX_EXTENSIONS = {'.docx'}
PDF_EXTENSIONS = {'.pdf'}
ALL_EXTENSIONS = MARKDOWN_EXTENSIONS | DOCX_EXTENSIONS | PDF_EXTENSIONS
ALL_EXTENSIONS_FOLDER = MARKDOWN_EXTENSIONS_FOLDER | DOCX_EXTENSIONS | PDF_EXTENSIONS


def get_conversion_direction(input_file: str, target_format: Optional[str] = None) -> str:
    """
    Determine conversion direction based on file extension and target format

    Args:
        input_file: Path to input file
        target_format: Target format ('pdf', 'docx', 'md') or None for default

    Returns:
        One of: 'md_to_docx', 'md_to_pdf', 'docx_to_md', 'docx_to_pdf',
                'pdf_to_md', 'pdf_to_docx'
    """
    suffix = Path(input_file).suffix.lower()

    if suffix in MARKDOWN_EXTENSIONS:
        if target_format == 'pdf':
            return 'md_to_pdf'
        return 'md_to_docx'  # Default
    elif suffix in DOCX_EXTENSIONS:
        if target_format == 'pdf':
            return 'docx_to_pdf'
        return 'docx_to_md'  # Default
    elif suffix in PDF_EXTENSIONS:
        if target_format == 'docx':
            return 'pdf_to_docx'
        return 'pdf_to_md'  # Default
    else:
        raise ValueError(f"Unsupported file type: {suffix}. Use .md, .markdown, .txt, .docx, or .pdf")


def is_convertible_file(file_path: Path, include_txt: bool = True) -> bool:
    """Check if a file can be converted

    Args:
        file_path: Path to check
        include_txt: Whether to include .txt files (False for folder scanning)
    """
    extensions = ALL_EXTENSIONS if include_txt else ALL_EXTENSIONS_FOLDER
    return file_path.suffix.lower() in extensions


def collect_files_from_folder(folder_path: Path, recursive: bool = False) -> List[Path]:
    """
    Collect all convertible files from a folder

    Note: .txt files are excluded from folder scanning to avoid converting
    non-markdown text files. Use explicit file paths to convert .txt files.

    Args:
        folder_path: Path to the folder
        recursive: Whether to search subdirectories

    Returns:
        List of file paths
    """
    files = []

    if recursive:
        for root, dirs, filenames in os.walk(folder_path):
            # Skip output directories we create
            dirs[:] = [d for d in dirs if d not in ('MD', 'DOCX', 'PDF')]
            for filename in filenames:
                file_path = Path(root) / filename
                if is_convertible_file(file_path, include_txt=False):
                    files.append(file_path)
    else:
        for item in folder_path.iterdir():
            if item.is_file() and is_convertible_file(item, include_txt=False):
                files.append(item)

    return sorted(files)


def get_converter_for_direction(direction: str, input_file: str, output_file: str = None):
    """
    Get the appropriate converter for a given direction

    Args:
        direction: Conversion direction string
        input_file: Path to input file
        output_file: Path to output file (optional)

    Returns:
        Converter instance
    """
    converters = {
        'md_to_docx': MarkdownToDocxConverter,
        'md_to_pdf': MarkdownToPdfConverter,
        'docx_to_md': DocxToMarkdownConverter,
        'docx_to_pdf': DocxToPdfConverter,
        'pdf_to_md': PdfToMarkdownConverter,
        'pdf_to_docx': PdfToDocxConverter,
    }

    converter_class = converters.get(direction)
    if not converter_class:
        raise ValueError(f"Unknown conversion direction: {direction}")

    return converter_class(input_file, output_file)


def convert_file(input_file: str, output_file: str = None, verbose: bool = False,
                 target_format: Optional[str] = None) -> str:
    """
    Convert a file (auto-detects direction)

    Args:
        input_file: Path to input file
        output_file: Path to output file (optional)
        verbose: Enable verbose output
        target_format: Target format ('pdf', 'docx', 'md') or None for default

    Returns:
        Path to the output file
    """
    direction = get_conversion_direction(input_file, target_format)
    converter = get_converter_for_direction(direction, input_file, output_file)
    return converter.convert()


def convert_folder(folder_path: Path, recursive: bool = False, verbose: bool = False,
                   target_format: Optional[str] = None) -> Tuple[int, int, int]:
    """
    Convert all files in a folder, organizing output into MD/, DOCX/, and PDF/ subfolders

    Args:
        folder_path: Path to the folder to convert
        recursive: Whether to process subdirectories
        verbose: Enable verbose output
        target_format: Target format ('pdf', 'docx', 'md') or None for default

    Returns:
        Tuple of (success_count, error_count, skipped_count)
    """
    # Collect files
    files = collect_files_from_folder(folder_path, recursive)

    if not files:
        print(f"No convertible files found in: {folder_path}")
        return 0, 0, 0

    # Create output directories
    md_output_dir = folder_path / 'MD'
    docx_output_dir = folder_path / 'DOCX'
    pdf_output_dir = folder_path / 'PDF'

    md_output_dir.mkdir(exist_ok=True)
    docx_output_dir.mkdir(exist_ok=True)
    pdf_output_dir.mkdir(exist_ok=True)

    if verbose:
        print(f"Created output directories:")
        print(f"  - {md_output_dir}")
        print(f"  - {docx_output_dir}")
        print(f"  - {pdf_output_dir}")

    success_count = 0
    error_count = 0
    skipped_count = 0

    # Count files by type for progress
    md_files = [f for f in files if f.suffix.lower() in MARKDOWN_EXTENSIONS_FOLDER]
    docx_files = [f for f in files if f.suffix.lower() in DOCX_EXTENSIONS]
    pdf_files = [f for f in files if f.suffix.lower() in PDF_EXTENSIONS]

    print(f"\nFound {len(files)} convertible files:")
    print(f"  - {len(md_files)} Markdown files")
    print(f"  - {len(docx_files)} Word files")
    print(f"  - {len(pdf_files)} PDF files")
    print()

    for i, file_path in enumerate(files, 1):
        try:
            direction = get_conversion_direction(str(file_path), target_format)

            # Determine output path based on direction
            output_info = {
                'md_to_docx': (docx_output_dir, '.docx', 'MD -> DOCX'),
                'md_to_pdf': (pdf_output_dir, '.pdf', 'MD -> PDF'),
                'docx_to_md': (md_output_dir, '.md', 'DOCX -> MD'),
                'docx_to_pdf': (pdf_output_dir, '.pdf', 'DOCX -> PDF'),
                'pdf_to_md': (md_output_dir, '.md', 'PDF -> MD'),
                'pdf_to_docx': (docx_output_dir, '.docx', 'PDF -> DOCX'),
            }

            output_dir, suffix, direction_str = output_info[direction]
            output_filename = file_path.stem + suffix
            output_path = output_dir / output_filename

            # Check if output already exists
            if output_path.exists():
                if verbose:
                    print(f"[{i}/{len(files)}] [SKIP] {file_path.name} (output exists)")
                skipped_count += 1
                continue

            if verbose:
                print(f"[{i}/{len(files)}] Converting: {file_path.name}")

            # Convert using the appropriate converter
            converter = get_converter_for_direction(direction, str(file_path), str(output_path))
            converter.convert()

            print(f"[{i}/{len(files)}] [OK] {direction_str}: {file_path.name} -> {output_path.name}")
            success_count += 1

        except Exception as e:
            print(f"[{i}/{len(files)}] [ERROR] {file_path.name}: {str(e)}", file=sys.stderr)
            error_count += 1
            if verbose:
                import traceback
                traceback.print_exc()

    return success_count, error_count, skipped_count


def main():
    """Main entry point for the unified converter"""
    parser = argparse.ArgumentParser(
        description='Convert between Markdown, Word, and PDF documents (auto-detects direction)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Conversion Direction (default):
  .md, .markdown, .txt  ->  .docx  (Markdown to Word)
  .docx                 ->  .md    (Word to Markdown)
  .pdf                  ->  .md    (PDF to Markdown)

With --to-pdf flag:
  .md, .markdown, .txt  ->  .pdf   (Markdown to PDF)
  .docx                 ->  .pdf   (Word to PDF)

With --to-docx flag:
  .pdf                  ->  .docx  (PDF to Word)

Folder Conversion:
  When a folder is provided, all convertible files are processed.
  Output files are organized into subfolders:
    - MD/    contains converted .md files
    - DOCX/  contains converted .docx files
    - PDF/   contains converted .pdf files

Examples:
  %(prog)s document.md              # Creates document.docx
  %(prog)s document.docx            # Creates document.md
  %(prog)s document.pdf             # Creates document.md
  %(prog)s document.md --to-pdf     # Creates document.pdf
  %(prog)s document.docx --to-pdf   # Creates document.pdf
  %(prog)s document.pdf --to-docx   # Creates document.docx
  %(prog)s report.md -o final.docx  # Specify output name
  %(prog)s *.md                     # Batch convert all markdown files
  %(prog)s ./my_folder              # Convert all files in folder
  %(prog)s ./my_folder --to-pdf     # Convert folder to PDF
  %(prog)s ./my_folder -r           # Convert folder recursively
        """
    )

    parser.add_argument(
        'input_paths',
        nargs='+',
        help='Input file(s) or folder(s) to convert'
    )

    parser.add_argument(
        '-o', '--output',
        help='Output file path (only valid with single input file, not folders)'
    )

    parser.add_argument(
        '-r', '--recursive',
        action='store_true',
        help='Recursively process subdirectories (for folder input)'
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose output'
    )

    parser.add_argument(
        '--to-docx',
        action='store_true',
        help='Force conversion to DOCX'
    )

    parser.add_argument(
        '--to-md',
        action='store_true',
        help='Force conversion to Markdown'
    )

    parser.add_argument(
        '--to-pdf',
        action='store_true',
        help='Force conversion to PDF'
    )

    parser.add_argument(
        '--overwrite',
        action='store_true',
        help='Overwrite existing output files (for folder conversion)'
    )

    args = parser.parse_args()

    # Check if any input is a directory
    has_directory = any(Path(p).is_dir() for p in args.input_paths)

    # Validate arguments
    if args.output and (len(args.input_paths) > 1 or has_directory):
        print("Error: -o/--output option can only be used with a single input file (not folders)")
        sys.exit(1)

    # Check for conflicting format flags
    format_flags = sum([args.to_docx, args.to_md, args.to_pdf])
    if format_flags > 1:
        print("Error: Cannot specify multiple output format flags (--to-docx, --to-md, --to-pdf)")
        sys.exit(1)

    # Determine target format from flags
    target_format = None
    if args.to_pdf:
        target_format = 'pdf'
    elif args.to_docx:
        target_format = 'docx'
    elif args.to_md:
        target_format = 'md'

    total_success = 0
    total_errors = 0
    total_skipped = 0

    for input_path in args.input_paths:
        path = Path(input_path)

        if path.is_dir():
            # Folder conversion mode
            print(f"\n{'='*60}")
            print(f"Processing folder: {path}")
            print('='*60)

            success, errors, skipped = convert_folder(
                path,
                recursive=args.recursive,
                verbose=args.verbose,
                target_format=target_format
            )

            total_success += success
            total_errors += errors
            total_skipped += skipped

        elif path.is_file():
            # Single file conversion
            try:
                if args.verbose:
                    print(f"Processing: {input_path}")

                # Determine conversion direction
                try:
                    direction = get_conversion_direction(input_path, target_format)
                except ValueError as e:
                    print(f"[SKIP] {input_path}: {e}")
                    total_skipped += 1
                    continue

                # Direction display strings
                direction_strings = {
                    'md_to_docx': 'MD -> DOCX',
                    'md_to_pdf': 'MD -> PDF',
                    'docx_to_md': 'DOCX -> MD',
                    'docx_to_pdf': 'DOCX -> PDF',
                    'pdf_to_md': 'PDF -> MD',
                    'pdf_to_docx': 'PDF -> DOCX',
                }
                direction_str = direction_strings.get(direction, direction)

                # Create appropriate converter
                converter = get_converter_for_direction(
                    direction,
                    input_path,
                    args.output if args.output else None
                )

                output_path = converter.convert()

                print(f"[OK] {direction_str}: {input_path} -> {output_path}")
                total_success += 1

            except FileNotFoundError:
                print(f"[ERROR] File not found: {input_path}", file=sys.stderr)
                total_errors += 1
                if args.verbose:
                    import traceback
                    traceback.print_exc()

            except Exception as e:
                print(f"[ERROR] {input_path}: {str(e)}", file=sys.stderr)
                total_errors += 1
                if args.verbose:
                    import traceback
                    traceback.print_exc()
        else:
            print(f"[ERROR] Path not found: {input_path}", file=sys.stderr)
            total_errors += 1

    # Print summary
    total_processed = total_success + total_errors + total_skipped
    if total_processed > 1:
        print(f"\n{'='*60}")
        print("SUMMARY")
        print('='*60)
        print(f"  Converted: {total_success}")
        print(f"  Errors:    {total_errors}")
        print(f"  Skipped:   {total_skipped}")
        print(f"  Total:     {total_processed}")

    sys.exit(0 if total_errors == 0 else 1)


if __name__ == '__main__':
    main()
