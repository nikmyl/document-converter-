#!/usr/bin/env python3
"""
Markdown to PDF Converter
Converts Markdown (.md) files to PDF documents using ReportLab
"""

import re
import sys
import argparse
from pathlib import Path
from typing import Optional, List, Tuple

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch, cm
    from reportlab.lib.colors import HexColor, black, blue
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Preformatted,
        Table, TableStyle, ListFlowable, ListItem, HRFlowable
    )
    from reportlab.lib import colors
except ImportError:
    print("Error: reportlab is required. Install it with: pip install reportlab")
    sys.exit(1)


class MarkdownToPdfConverter:
    """Convert Markdown files to PDF documents"""

    # Supported input extensions
    SUPPORTED_EXTENSIONS = {'.md', '.markdown', '.txt'}

    def __init__(self, input_file: str, output_file: Optional[str] = None):
        """
        Initialize the converter

        Args:
            input_file: Path to the input .md, .markdown, or .txt file
            output_file: Path to the output .pdf file (optional)
        """
        self.input_file = Path(input_file)

        if not self.input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")

        if self.input_file.suffix.lower() not in self.SUPPORTED_EXTENSIONS:
            raise ValueError("Input file must be a .md, .markdown, or .txt file")

        if output_file:
            self.output_file = Path(output_file)
        else:
            self.output_file = self.input_file.with_suffix('.pdf')

        self._setup_styles()

    def _setup_styles(self):
        """Set up document styles"""
        self.styles = getSampleStyleSheet()

        # Heading styles
        self.styles.add(ParagraphStyle(
            name='Heading1Custom',
            parent=self.styles['Heading1'],
            fontSize=24,
            spaceAfter=12,
            spaceBefore=18,
            textColor=HexColor('#2c3e50'),
            borderPadding=(0, 0, 6, 0),
            borderWidth=2,
            borderColor=HexColor('#3498db'),
        ))

        self.styles.add(ParagraphStyle(
            name='Heading2Custom',
            parent=self.styles['Heading2'],
            fontSize=20,
            spaceAfter=10,
            spaceBefore=14,
            textColor=HexColor('#2c3e50'),
        ))

        self.styles.add(ParagraphStyle(
            name='Heading3Custom',
            parent=self.styles['Heading3'],
            fontSize=16,
            spaceAfter=8,
            spaceBefore=12,
            textColor=HexColor('#2c3e50'),
        ))

        self.styles.add(ParagraphStyle(
            name='Heading4Custom',
            parent=self.styles['Heading4'],
            fontSize=14,
            spaceAfter=6,
            spaceBefore=10,
            textColor=HexColor('#2c3e50'),
        ))

        # Body text
        self.styles.add(ParagraphStyle(
            name='BodyCustom',
            parent=self.styles['Normal'],
            fontSize=11,
            leading=16,
            spaceAfter=8,
            alignment=TA_JUSTIFY,
        ))

        # Code block style
        self.styles.add(ParagraphStyle(
            name='CodeBlock',
            fontName='Courier',
            fontSize=9,
            leading=12,
            leftIndent=20,
            rightIndent=20,
            spaceBefore=8,
            spaceAfter=8,
            backColor=HexColor('#f8f8f8'),
        ))

        # Blockquote style
        self.styles.add(ParagraphStyle(
            name='BlockQuote',
            parent=self.styles['Normal'],
            fontSize=11,
            leading=14,
            leftIndent=30,
            rightIndent=20,
            spaceBefore=8,
            spaceAfter=8,
            textColor=HexColor('#555555'),
            fontName='Times-Italic',
            borderPadding=(8, 8, 8, 8),
            borderWidth=0,
            borderColor=HexColor('#3498db'),
        ))

        # List item style
        self.styles.add(ParagraphStyle(
            name='ListItem',
            parent=self.styles['Normal'],
            fontSize=11,
            leading=14,
            leftIndent=20,
            spaceAfter=4,
        ))

    def convert(self) -> str:
        """
        Convert the Markdown file to PDF

        Returns:
            Path to the output file
        """
        # Read the markdown content
        with open(self.input_file, 'r', encoding='utf-8') as f:
            md_content = f.read()

        # Create PDF document
        doc = SimpleDocTemplate(
            str(self.output_file),
            pagesize=A4,
            rightMargin=2*cm,
            leftMargin=2*cm,
            topMargin=2.5*cm,
            bottomMargin=2.5*cm
        )

        # Build the story (list of flowables)
        story = self._process_markdown(md_content)

        # Build PDF
        doc.build(story)

        return str(self.output_file)

    def _process_markdown(self, content: str) -> List:
        """Process markdown content and return list of flowables"""
        story = []
        lines = content.split('\n')
        i = 0
        in_code_block = False
        code_block_lines = []
        in_table = False
        table_rows = []

        while i < len(lines):
            line = lines[i]

            # Handle code blocks
            if line.strip().startswith('```'):
                if not in_code_block:
                    in_code_block = True
                    code_block_lines = []
                else:
                    # End of code block
                    if code_block_lines:
                        code_text = '\n'.join(code_block_lines)
                        story.append(Preformatted(code_text, self.styles['CodeBlock']))
                        story.append(Spacer(1, 8))
                    in_code_block = False
                    code_block_lines = []
                i += 1
                continue

            if in_code_block:
                code_block_lines.append(line)
                i += 1
                continue

            # Handle tables
            if '|' in line and line.strip().startswith('|'):
                if not in_table:
                    in_table = True
                    table_rows = []
                table_rows.append(line)
                i += 1
                continue
            elif in_table:
                # End of table
                if table_rows:
                    table_flowable = self._create_table(table_rows)
                    if table_flowable:
                        story.append(table_flowable)
                        story.append(Spacer(1, 12))
                in_table = False
                table_rows = []
                # Don't increment i, process current line

            # Handle headers
            header_match = re.match(r'^(#{1,6})\s+(.+)$', line)
            if header_match:
                level = len(header_match.group(1))
                text = self._process_inline_formatting(header_match.group(2))
                style_name = f'Heading{min(level, 4)}Custom'
                if style_name not in [s.name for s in self.styles.byName.values()]:
                    style_name = 'Heading4Custom'
                story.append(Paragraph(text, self.styles[style_name]))
                i += 1
                continue

            # Handle horizontal rules
            if re.match(r'^(\*{3,}|-{3,}|_{3,})$', line.strip()):
                story.append(HRFlowable(width="100%", thickness=1, color=HexColor('#bdc3c7')))
                story.append(Spacer(1, 12))
                i += 1
                continue

            # Handle unordered lists
            list_match = re.match(r'^[\*\-\+]\s+(.+)$', line)
            if list_match:
                items = []
                while i < len(lines):
                    lm = re.match(r'^[\*\-\+]\s+(.+)$', lines[i])
                    if lm:
                        text = self._process_inline_formatting(lm.group(1))
                        items.append(ListItem(Paragraph(text, self.styles['ListItem']), leftIndent=20, bulletColor=black))
                        i += 1
                    else:
                        break
                if items:
                    story.append(ListFlowable(items, bulletType='bullet', start=None))
                    story.append(Spacer(1, 8))
                continue

            # Handle ordered lists
            ordered_match = re.match(r'^(\d+)\.\s+(.+)$', line)
            if ordered_match:
                items = []
                while i < len(lines):
                    om = re.match(r'^(\d+)\.\s+(.+)$', lines[i])
                    if om:
                        text = self._process_inline_formatting(om.group(2))
                        items.append(ListItem(Paragraph(text, self.styles['ListItem']), leftIndent=20))
                        i += 1
                    else:
                        break
                if items:
                    story.append(ListFlowable(items, bulletType='1', start=1))
                    story.append(Spacer(1, 8))
                continue

            # Handle blockquotes
            if line.strip().startswith('>'):
                quote_lines = []
                while i < len(lines) and lines[i].strip().startswith('>'):
                    quote_text = lines[i].strip()[1:].strip()
                    quote_lines.append(quote_text)
                    i += 1
                if quote_lines:
                    quote_content = '<br/>'.join(quote_lines)
                    story.append(Paragraph(quote_content, self.styles['BlockQuote']))
                    story.append(Spacer(1, 8))
                continue

            # Handle empty lines
            if not line.strip():
                story.append(Spacer(1, 6))
                i += 1
                continue

            # Handle regular paragraphs
            text = self._process_inline_formatting(line)
            if text.strip():
                story.append(Paragraph(text, self.styles['BodyCustom']))
            i += 1

        # Handle any remaining table
        if in_table and table_rows:
            table_flowable = self._create_table(table_rows)
            if table_flowable:
                story.append(table_flowable)

        return story

    def _process_inline_formatting(self, text: str) -> str:
        """Process inline markdown formatting and return HTML-like markup for ReportLab"""
        # Escape special characters for ReportLab
        text = text.replace('&', '&amp;')
        text = text.replace('<', '&lt;')
        text = text.replace('>', '&gt;')

        # Inline code: `code` -> <font face="Courier" color="#c7254e">code</font>
        text = re.sub(r'`([^`]+)`', r'<font face="Courier" color="#c7254e">\1</font>', text)

        # Links first (before other processing can interfere)
        text = re.sub(
            r'\[([^\]]+)\]\(([^)]+)\)',
            r'<link href="\2"><u><font color="blue">\1</font></u></link>',
            text
        )

        # Remove strikethrough markers (not supported in ReportLab Paragraph)
        text = re.sub(r'~~([^~]+)~~', r'\1', text)

        # Mixed bold+italic patterns: **_text_** or _**text**_ or *__text__* etc.
        text = re.sub(r'\*\*_([^_*]+)_\*\*', r'<b><i>\1</i></b>', text)
        text = re.sub(r'_\*\*([^_*]+)\*\*_', r'<i><b>\1</b></i>', text)
        text = re.sub(r'\*__([^_*]+)__\*', r'<i><b>\1</b></i>', text)
        text = re.sub(r'__\*([^_*]+)\*__', r'<b><i>\1</i></b>', text)

        # Bold and italic: ***text*** or ___text___
        text = re.sub(r'\*\*\*([^*]+)\*\*\*', r'<b><i>\1</i></b>', text)
        text = re.sub(r'___([^_]+)___', r'<b><i>\1</i></b>', text)

        # Bold: **text** or __text__
        text = re.sub(r'\*\*([^*]+)\*\*', r'<b>\1</b>', text)
        text = re.sub(r'__([^_]+)__', r'<b>\1</b>', text)

        # Italic: *text* or _text_
        text = re.sub(r'(?<!\*)\*([^*]+)\*(?!\*)', r'<i>\1</i>', text)
        text = re.sub(r'(?<!_)_([^_]+)_(?!_)', r'<i>\1</i>', text)

        # Clean up any remaining isolated markers
        text = re.sub(r'(?<!\w)\*+(?!\w)', '', text)
        text = re.sub(r'(?<!\w)_+(?!\w)', '', text)

        return text

    def _create_table(self, table_rows: List[str]) -> Optional[Table]:
        """Create a ReportLab table from markdown table rows"""
        if not table_rows:
            return None

        # Parse table rows
        parsed_rows = []
        for row in table_rows:
            cells = [cell.strip() for cell in row.split('|')]
            if cells and cells[0] == '':
                cells = cells[1:]
            if cells and cells[-1] == '':
                cells = cells[:-1]

            # Skip separator row
            if cells and all(set(cell.replace('-', '').replace(':', '').replace(' ', '')) == set() for cell in cells):
                continue

            if cells:
                # Process inline formatting in cells
                cells = [self._process_inline_formatting(cell) for cell in cells]
                parsed_rows.append(cells)

        if not parsed_rows:
            return None

        # Determine column count
        num_cols = max(len(row) for row in parsed_rows)

        # Normalize rows
        for row in parsed_rows:
            while len(row) < num_cols:
                row.append('')

        # Convert to Paragraph objects for proper formatting
        table_data = []
        for i, row in enumerate(parsed_rows):
            table_row = []
            for cell in row:
                style = self.styles['BodyCustom']
                p = Paragraph(cell, style)
                table_row.append(p)
            table_data.append(table_row)

        # Create table
        table = Table(table_data, repeatRows=1)

        # Style the table
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), HexColor('#3498db')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
            ('TOPPADDING', (0, 0), (-1, 0), 10),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, HexColor('#f9f9f9')]),
            ('GRID', (0, 0), (-1, -1), 1, HexColor('#dddddd')),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('TOPPADDING', (0, 1), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ])
        table.setStyle(table_style)

        return table


def main():
    """Main entry point for the script"""
    parser = argparse.ArgumentParser(
        description='Convert Markdown files to PDF documents',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s input.md
  %(prog)s input.md -o output.pdf
  %(prog)s *.md
        """
    )

    parser.add_argument(
        'input_files',
        nargs='+',
        help='Input Markdown file(s) to convert'
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

            converter = MarkdownToPdfConverter(
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
