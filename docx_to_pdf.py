#!/usr/bin/env python3
"""
DOCX to PDF Converter
Converts Word documents (.docx) to PDF files by chaining through Markdown
"""

import sys
import argparse
import tempfile
from pathlib import Path
from typing import Optional

# Import existing converters
try:
    from docx_to_md import DocxToMarkdownConverter
    from md_to_pdf import MarkdownToPdfConverter
except ImportError as e:
    print(f"Error: Required converter modules not found: {e}")
    sys.exit(1)


class DocxToPdfConverter:
    """Convert Word documents to PDF files"""

    def __init__(self, input_file: str, output_file: Optional[str] = None):
        """
        Initialize the converter

        Args:
            input_file: Path to the input .docx file
            output_file: Path to the output .pdf file (optional)
        """
        self.input_file = Path(input_file)

        if not self.input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")

        if not self.input_file.suffix.lower() == '.docx':
            raise ValueError("Input file must be a .docx (Word) file")

        if output_file:
            self.output_file = Path(output_file)
        else:
            self.output_file = self.input_file.with_suffix('.pdf')

    def convert(self) -> str:
        """
        Convert the DOCX file to PDF

        Strategy: DOCX -> MD (temp) -> PDF
        This ensures consistent output by reusing existing converters.

        Returns:
            Path to the output file
        """
        # Create a temporary markdown file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as tmp:
            temp_md_path = tmp.name

        try:
            # Step 1: Convert DOCX to Markdown
            docx_to_md = DocxToMarkdownConverter(
                str(self.input_file),
                temp_md_path
            )
            docx_to_md.convert()

            # Step 2: Convert Markdown to PDF
            md_to_pdf = MarkdownToPdfConverter(
                temp_md_path,
                str(self.output_file)
            )
            md_to_pdf.convert()

            return str(self.output_file)

        finally:
            # Clean up temporary file
            try:
                Path(temp_md_path).unlink()
            except Exception:
                pass


def main():
    """Main entry point for the script"""
    parser = argparse.ArgumentParser(
        description='Convert Word documents (.docx) to PDF files',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s document.docx
  %(prog)s document.docx -o output.pdf
  %(prog)s *.docx
        """
    )

    parser.add_argument(
        'input_files',
        nargs='+',
        help='Input Word file(s) to convert'
    )

    parser.add_argument(
        '-o', '--output',
        help='Output file path (only valid with single input file)'
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose output'
    )

    args = parser.parse_args()

    # Validate arguments
    if args.output and len(args.input_files) > 1:
        print("Error: -o/--output option can only be used with a single input file")
        sys.exit(1)

    # Process each file
    for input_file in args.input_files:
        try:
            if args.verbose:
                print(f"Converting: {input_file}")

            converter = DocxToPdfConverter(
                input_file,
                args.output if args.output else None
            )
            output_path = converter.convert()

            print(f"[OK] Converted: {input_file} -> {output_path}")

        except Exception as e:
            print(f"[ERROR] Error converting {input_file}: {str(e)}", file=sys.stderr)
            if args.verbose:
                import traceback
                traceback.print_exc()


if __name__ == '__main__':
    main()
