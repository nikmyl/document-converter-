#!/usr/bin/env python3
"""
PDF to Markdown Converter
Converts PDF documents to Markdown (.md) files using pdfplumber and pypdf
"""

import re
import sys
import argparse
from pathlib import Path
from typing import Optional, List, Dict, Tuple, Any
from dataclasses import dataclass

try:
    import pdfplumber
except ImportError:
    print("Error: pdfplumber is required. Install it with: pip install pdfplumber")
    sys.exit(1)

try:
    from pypdf import PdfReader
except ImportError:
    print("Error: pypdf is required. Install it with: pip install pypdf")
    sys.exit(1)


@dataclass
class TextBlock:
    """Represents a text block with metadata"""
    text: str
    page_num: int
    top: float
    bottom: float
    font_size: float
    is_bold: bool
    is_italic: bool


class PdfToMarkdownConverter:
    """Convert PDF documents to Markdown files"""

    def __init__(self, input_file: str, output_file: Optional[str] = None):
        """
        Initialize the converter

        Args:
            input_file: Path to the input .pdf file
            output_file: Path to the output .md file (optional)
        """
        self.input_file = Path(input_file)

        if not self.input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")

        if not self.input_file.suffix.lower() == '.pdf':
            raise ValueError("Input file must be a .pdf file")

        if output_file:
            self.output_file = Path(output_file)
        else:
            self.output_file = self.input_file.with_suffix('.md')

        self.markdown_lines: List[str] = []
        self._font_sizes: List[float] = []
        self._base_font_size: float = 12.0

    def convert(self) -> str:
        """
        Convert the PDF file to Markdown

        Returns:
            Path to the output file
        """
        # Process the PDF
        self._process_pdf()

        # Write the markdown content
        markdown_content = '\n'.join(self.markdown_lines)

        # Clean up excessive blank lines
        markdown_content = re.sub(r'\n{3,}', '\n\n', markdown_content)

        with open(self.output_file, 'w', encoding='utf-8') as f:
            f.write(markdown_content)

        return str(self.output_file)

    def _process_pdf(self):
        """Process the PDF document"""
        with pdfplumber.open(str(self.input_file)) as pdf:
            # First pass: collect font sizes to determine base size
            self._analyze_font_sizes(pdf)

            # Second pass: extract content
            for page_num, page in enumerate(pdf.pages):
                self._process_page(page, page_num)

    def _analyze_font_sizes(self, pdf):
        """Analyze font sizes across the document to determine heading thresholds"""
        font_sizes = []

        for page in pdf.pages:
            chars = page.chars
            for char in chars:
                if char.get('size'):
                    font_sizes.append(char['size'])

        if font_sizes:
            # Use the most common font size as base
            from collections import Counter
            size_counts = Counter(round(s, 1) for s in font_sizes)
            self._base_font_size = size_counts.most_common(1)[0][0]
            self._font_sizes = sorted(set(round(s, 1) for s in font_sizes), reverse=True)

    def _process_page(self, page, page_num: int):
        """Process a single PDF page"""
        # Extract tables first (to know which regions to skip for text)
        tables = page.find_tables()
        table_bboxes = [table.bbox for table in tables]

        # Extract and process tables
        for table in tables:
            table_data = table.extract()
            if table_data:
                self._add_table(table_data)

        # Extract text, avoiding table regions
        text_content = self._extract_text_avoiding_tables(page, table_bboxes)

        # Process text content
        self._process_text_content(text_content, page)

    def _extract_text_avoiding_tables(self, page, table_bboxes: List[Tuple]) -> str:
        """Extract text from page while avoiding table regions"""
        if not table_bboxes:
            return page.extract_text() or ""

        # Create a filtered page that excludes table areas
        # We'll extract text and filter out lines that fall within table regions
        chars = page.chars
        filtered_chars = []

        for char in chars:
            char_y = char['top']
            in_table = False

            for bbox in table_bboxes:
                # bbox is (x0, top, x1, bottom)
                if bbox[1] <= char_y <= bbox[3]:
                    in_table = True
                    break

            if not in_table:
                filtered_chars.append(char)

        # Reconstruct text from filtered characters
        if not filtered_chars:
            return ""

        # Group chars by line (similar y position)
        lines = []
        current_line = []
        current_y = None
        y_tolerance = 3

        sorted_chars = sorted(filtered_chars, key=lambda c: (c['top'], c['x0']))

        for char in sorted_chars:
            if current_y is None or abs(char['top'] - current_y) <= y_tolerance:
                current_line.append(char)
                current_y = char['top']
            else:
                if current_line:
                    lines.append(current_line)
                current_line = [char]
                current_y = char['top']

        if current_line:
            lines.append(current_line)

        # Convert lines to text
        text_lines = []
        for line in lines:
            sorted_line = sorted(line, key=lambda c: c['x0'])
            line_text = ''.join(c.get('text', '') for c in sorted_line)
            text_lines.append(line_text)

        return '\n'.join(text_lines)

    def _process_text_content(self, text: str, page):
        """Process extracted text and convert to markdown"""
        if not text:
            return

        lines = text.split('\n')
        chars = page.chars

        for line in lines:
            line = line.strip()
            if not line:
                if self.markdown_lines and self.markdown_lines[-1] != '':
                    self.markdown_lines.append('')
                continue

            # Detect line characteristics
            line_chars = self._get_chars_for_line(chars, line)
            avg_font_size = self._get_average_font_size(line_chars)
            is_bold = self._is_bold(line_chars)

            # Check for heading
            heading_level = self._detect_heading_level(avg_font_size, is_bold)
            if heading_level:
                self.markdown_lines.append('')
                self.markdown_lines.append(f"{'#' * heading_level} {line}")
                self.markdown_lines.append('')
                continue

            # Check for bullet list
            bullet_match = re.match(r'^[\u2022\u2023\u25E6\u2043\u2219•●○◦-]\s*(.+)$', line)
            if bullet_match:
                self.markdown_lines.append(f"- {bullet_match.group(1)}")
                continue

            # Check for numbered list (support 10+ item lists)
            numbered_match = re.match(r'^(\d+)[.)\]]\s*(.+)$', line)
            if numbered_match:
                num = numbered_match.group(1)
                text = numbered_match.group(2)
                self.markdown_lines.append(f"{num}. {text}")
                continue

            # Check for horizontal rule
            if re.match(r'^[-_=]{3,}$', line.replace(' ', '')):
                self.markdown_lines.append('')
                self.markdown_lines.append('---')
                self.markdown_lines.append('')
                continue

            # Regular paragraph with inline formatting
            formatted_line = self._apply_inline_formatting(line, line_chars)
            self.markdown_lines.append(formatted_line)

    def _get_chars_for_line(self, chars: List[Dict], line_text: str) -> List[Dict]:
        """Get character objects that correspond to a line of text"""
        # Simple matching - find chars that form this line
        matching_chars = []
        line_lower = line_text.lower().replace(' ', '')

        for char in chars:
            char_text = char.get('text', '')
            if char_text and char_text.lower() in line_lower:
                matching_chars.append(char)

        return matching_chars

    def _get_average_font_size(self, chars: List[Dict]) -> float:
        """Calculate average font size for a set of characters"""
        if not chars:
            return self._base_font_size

        sizes = [c.get('size', self._base_font_size) for c in chars if c.get('size')]
        return sum(sizes) / len(sizes) if sizes else self._base_font_size

    def _is_bold(self, chars: List[Dict]) -> bool:
        """Check if characters are bold"""
        if not chars:
            return False

        bold_count = 0
        for char in chars:
            fontname = char.get('fontname', '').lower()
            if 'bold' in fontname or 'heavy' in fontname or 'black' in fontname:
                bold_count += 1

        return bold_count > len(chars) / 2

    def _is_italic(self, chars: List[Dict]) -> bool:
        """Check if characters are italic"""
        if not chars:
            return False

        italic_count = 0
        for char in chars:
            fontname = char.get('fontname', '').lower()
            if 'italic' in fontname or 'oblique' in fontname:
                italic_count += 1

        return italic_count > len(chars) / 2

    def _detect_heading_level(self, font_size: float, is_bold: bool) -> Optional[int]:
        """
        Determine heading level from font characteristics

        Uses adaptive thresholds based on the document's font size distribution
        rather than fixed magic numbers.
        """
        if not self._font_sizes or font_size <= self._base_font_size:
            return None

        # Calculate size ratio
        ratio = font_size / self._base_font_size

        # Get the range of font sizes in the document
        max_font_size = max(self._font_sizes) if self._font_sizes else self._base_font_size
        font_range = max_font_size - self._base_font_size

        if font_range <= 0:
            return None

        # Calculate relative position in the font size range
        relative_size = (font_size - self._base_font_size) / font_range

        # Map relative position to heading levels
        # This adapts to documents with different font size distributions
        if relative_size >= 0.85:  # Top 15% of font range
            return 1
        elif relative_size >= 0.65:  # 65-85%
            return 2
        elif relative_size >= 0.45:  # 45-65%
            return 3
        elif relative_size >= 0.25:  # 25-45%
            return 4 if is_bold else None
        elif relative_size >= 0.10:  # 10-25%
            return 5 if is_bold else None
        elif is_bold and ratio > 1.05:  # Small but bold and larger than base
            return 6

        # Fallback to ratio-based detection for edge cases
        if ratio >= 1.8:
            return 1
        elif ratio >= 1.5:
            return 2
        elif ratio >= 1.3:
            return 3
        elif ratio >= 1.15 and is_bold:
            return 4

        return None

    def _apply_inline_formatting(self, line: str, chars: List[Dict]) -> str:
        """Apply inline markdown formatting based on character properties"""
        if not chars or not line:
            return line

        # Group characters by their formatting
        # This is a simplified approach - we check if majority of chars are bold/italic
        bold_count = 0
        italic_count = 0

        for char in chars:
            fontname = char.get('fontname', '').lower()
            if 'bold' in fontname or 'heavy' in fontname or 'black' in fontname:
                bold_count += 1
            if 'italic' in fontname or 'oblique' in fontname:
                italic_count += 1

        total_chars = len(chars)
        if total_chars == 0:
            return line

        # Apply formatting if majority of characters have that style
        is_mostly_bold = bold_count > total_chars * 0.7
        is_mostly_italic = italic_count > total_chars * 0.7

        if is_mostly_bold and is_mostly_italic:
            return f"***{line}***"
        elif is_mostly_bold:
            return f"**{line}**"
        elif is_mostly_italic:
            return f"*{line}*"

        return line

    def _add_table(self, table_data: List[List[Any]]):
        """Add a table to markdown output"""
        if not table_data or not table_data[0]:
            return

        self.markdown_lines.append('')

        # Filter out None values and empty rows
        cleaned_rows = []
        for row in table_data:
            if row:
                cleaned_row = [str(cell).strip() if cell else '' for cell in row]
                if any(cleaned_row):  # Only add non-empty rows
                    cleaned_rows.append(cleaned_row)

        if not cleaned_rows:
            return

        # Determine number of columns
        num_cols = max(len(row) for row in cleaned_rows)

        # Normalize rows to have same number of columns
        for row in cleaned_rows:
            while len(row) < num_cols:
                row.append('')

        # Create header row
        header = cleaned_rows[0]
        # Escape pipe characters
        header = [cell.replace('|', '\\|') for cell in header]
        self.markdown_lines.append('| ' + ' | '.join(header) + ' |')

        # Create separator row
        separator = '| ' + ' | '.join(['---'] * num_cols) + ' |'
        self.markdown_lines.append(separator)

        # Create data rows
        for row in cleaned_rows[1:]:
            row = [cell.replace('|', '\\|') for cell in row]
            self.markdown_lines.append('| ' + ' | '.join(row) + ' |')

        self.markdown_lines.append('')


def main():
    """Main entry point for the script"""
    parser = argparse.ArgumentParser(
        description='Convert PDF documents to Markdown files',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s document.pdf
  %(prog)s document.pdf -o output.md
  %(prog)s *.pdf
        """
    )

    parser.add_argument(
        'input_files',
        nargs='+',
        help='Input PDF file(s) to convert'
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

            converter = PdfToMarkdownConverter(
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
