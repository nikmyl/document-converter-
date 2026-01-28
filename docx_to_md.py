#!/usr/bin/env python3
"""
DOCX to Markdown Converter
Converts Microsoft Word documents (.docx) to Markdown (.md) files
"""

import re
import sys
from pathlib import Path
from typing import Optional, List
from xml.etree import ElementTree as ET

try:
    from docx import Document
    from docx.document import Document as DocumentType
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run
    from docx.table import Table, _Cell
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("Error: python-docx is required. Install it with: pip install python-docx")
    sys.exit(1)


class DocxToMarkdownConverter:
    """Convert Word documents to Markdown files"""

    # XML namespaces used in DOCX
    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    }

    def __init__(self, input_file: str, output_file: Optional[str] = None):
        """
        Initialize the converter

        Args:
            input_file: Path to the input .docx file
            output_file: Path to the output .md file (optional)
        """
        self.input_file = Path(input_file)

        if not self.input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")

        if not self.input_file.suffix.lower() == '.docx':
            raise ValueError("Input file must be a .docx (Word) file")

        if output_file:
            self.output_file = Path(output_file)
        else:
            self.output_file = self.input_file.with_suffix('.md')

        self.doc = Document(str(self.input_file))
        self.markdown_lines: List[str] = []
        self._list_state = {'in_list': False, 'list_type': None, 'item_count': 0}

    def convert(self) -> str:
        """
        Convert the DOCX file to Markdown

        Returns:
            Path to the output file
        """
        self._process_document()

        # Write the markdown content
        markdown_content = '\n'.join(self.markdown_lines)

        # Clean up excessive blank lines
        markdown_content = re.sub(r'\n{3,}', '\n\n', markdown_content)

        with open(self.output_file, 'w', encoding='utf-8') as f:
            f.write(markdown_content)

        return str(self.output_file)

    def _process_document(self):
        """Process the entire document"""
        for element in self.doc.element.body:
            if element.tag == qn('w:p'):
                # It's a paragraph
                paragraph = Paragraph(element, self.doc)
                self._process_paragraph(paragraph)
            elif element.tag == qn('w:tbl'):
                # It's a table
                table = Table(element, self.doc)
                self._process_table(table)

    def _process_paragraph(self, paragraph: Paragraph):
        """Process a single paragraph"""
        # Check if this is a list item
        is_list_item, list_type = self._is_list_item(paragraph)

        # Get the paragraph style
        style_name = paragraph.style.name if paragraph.style else 'Normal'

        # Handle headings
        if style_name.startswith('Heading'):
            self._end_list()
            level = self._get_heading_level(style_name)
            text = self._get_formatted_text(paragraph)
            if text.strip():
                self.markdown_lines.append('')
                self.markdown_lines.append(f"{'#' * level} {text.strip()}")
                self.markdown_lines.append('')
            return

        # Handle list items
        if is_list_item:
            text = self._get_formatted_text(paragraph)
            if text.strip():
                if list_type == 'bullet':
                    self.markdown_lines.append(f"- {text.strip()}")
                else:
                    self._list_state['item_count'] += 1
                    self.markdown_lines.append(f"{self._list_state['item_count']}. {text.strip()}")
                self._list_state['in_list'] = True
                self._list_state['list_type'] = list_type
            return

        # End any ongoing list
        if self._list_state['in_list']:
            self._end_list()

        # Handle blockquotes (Quote style)
        if style_name == 'Quote' or style_name == 'Intense Quote':
            text = self._get_formatted_text(paragraph)
            if text.strip():
                self.markdown_lines.append(f"> {text.strip()}")
            return

        # Handle code blocks (by font name)
        if self._is_code_block(paragraph):
            text = paragraph.text
            if text.strip():
                self.markdown_lines.append(f"```")
                self.markdown_lines.append(text)
                self.markdown_lines.append(f"```")
                self.markdown_lines.append('')
            return

        # Handle horizontal rules
        if self._is_horizontal_rule(paragraph):
            self.markdown_lines.append('')
            self.markdown_lines.append('---')
            self.markdown_lines.append('')
            return

        # Handle regular paragraphs
        text = self._get_formatted_text(paragraph)
        if text.strip():
            self.markdown_lines.append(text.strip())
            self.markdown_lines.append('')
        elif self.markdown_lines and self.markdown_lines[-1] != '':
            # Add blank line for empty paragraphs (but avoid multiple)
            self.markdown_lines.append('')

    def _end_list(self):
        """End the current list and reset state"""
        if self._list_state['in_list']:
            self.markdown_lines.append('')
            self._list_state = {'in_list': False, 'list_type': None, 'item_count': 0}

    def _is_list_item(self, paragraph: Paragraph) -> tuple:
        """Check if paragraph is a list item and determine list type"""
        # Check style name
        style_name = paragraph.style.name if paragraph.style else ''

        if 'List Bullet' in style_name:
            return True, 'bullet'
        if 'List Number' in style_name:
            return True, 'numbered'

        # Check for numbering in XML
        p_element = paragraph._element
        numPr = p_element.find('.//w:numPr', self.NAMESPACES)

        if numPr is not None:
            numId = numPr.find('w:numId', self.NAMESPACES)
            if numId is not None:
                # Try to determine if bullet or numbered
                # This is a simplified check
                ilvl = numPr.find('w:ilvl', self.NAMESPACES)
                if ilvl is not None:
                    # Check the numbering definition for type
                    # For simplicity, we'll check common patterns
                    if 'bullet' in style_name.lower():
                        return True, 'bullet'
                    return True, 'numbered'

        return False, None

    def _get_heading_level(self, style_name: str) -> int:
        """Extract heading level from style name"""
        match = re.search(r'Heading\s*(\d+)', style_name)
        if match:
            return min(int(match.group(1)), 6)
        return 1

    def _get_formatted_text(self, paragraph: Paragraph) -> str:
        """Get paragraph text with inline formatting"""
        result = []

        for run in paragraph.runs:
            text = run.text
            if not text:
                continue

            # Check for hyperlinks (handled separately)
            # Apply formatting
            if run.bold and run.italic:
                text = f"***{text}***"
            elif run.bold:
                text = f"**{text}**"
            elif run.italic:
                text = f"*{text}*"

            # Check for inline code (monospace font)
            if run.font.name and 'courier' in run.font.name.lower():
                text = f"`{run.text}`"

            result.append(text)

        # Handle hyperlinks
        formatted = ''.join(result)
        formatted = self._process_hyperlinks(paragraph, formatted)

        return formatted

    def _process_hyperlinks(self, paragraph: Paragraph, text: str) -> str:
        """Extract and format hyperlinks from paragraph"""
        # Get hyperlinks from the paragraph XML
        p_element = paragraph._element
        hyperlinks = p_element.findall('.//w:hyperlink', self.NAMESPACES)

        for hyperlink in hyperlinks:
            # Get the relationship ID
            r_id = hyperlink.get(qn('r:id'))
            if r_id:
                try:
                    # Get the target URL from relationships
                    rel = self.doc.part.rels.get(r_id)
                    if rel and hasattr(rel, 'target_ref'):
                        url = rel.target_ref
                        # Get the hyperlink text
                        link_text = ''.join([node.text or '' for node in hyperlink.iter() if node.text])
                        if link_text and url:
                            # Replace plain text with markdown link
                            text = text.replace(link_text, f"[{link_text}]({url})", 1)
                except Exception:
                    pass

        return text

    def _is_code_block(self, paragraph: Paragraph) -> bool:
        """
        Check if paragraph is a code block

        Only treat as code block if:
        1. Style explicitly indicates code (e.g., 'CodeBlock', 'Code')
        2. ALL runs have monospace font AND it's a multi-line or indented block

        This prevents regular monospace text from being treated as code blocks.
        """
        style_name = paragraph.style.name if paragraph.style else ''

        # Check for explicit CodeBlock style
        if 'code' in style_name.lower():
            return True

        # Check paragraph formatting that suggests code block
        # (e.g., left indent, specific style)
        if paragraph.paragraph_format.left_indent and paragraph.paragraph_format.left_indent.inches >= 0.3:
            # Has significant left indent - could be code block
            # Check if ALL runs have monospace font
            if paragraph.runs:
                all_monospace = all(
                    run.font.name and 'courier' in run.font.name.lower()
                    for run in paragraph.runs if run.text.strip()
                )
                if all_monospace:
                    return True

        return False

    def _is_horizontal_rule(self, paragraph: Paragraph) -> bool:
        """Check if paragraph represents a horizontal rule"""
        text = paragraph.text.strip()
        # Check for common horizontal rule patterns
        if re.match(r'^[-_]{3,}$', text.replace(' ', '')):
            return True
        if text == '_' * 50:  # Our converter creates this
            return True
        return False

    def _process_table(self, table: Table):
        """Convert a Word table to Markdown table"""
        if not table.rows:
            return

        self.markdown_lines.append('')

        rows_data = []
        for row in table.rows:
            row_cells = []
            for cell in row.cells:
                # Get cell text, handling multiple paragraphs
                cell_text = ' '.join(p.text.strip() for p in cell.paragraphs if p.text.strip())
                # Escape pipe characters in cell content
                cell_text = cell_text.replace('|', '\\|')
                row_cells.append(cell_text)
            rows_data.append(row_cells)

        if not rows_data:
            return

        # Determine number of columns
        num_cols = max(len(row) for row in rows_data)

        # Normalize rows to have same number of columns
        for row in rows_data:
            while len(row) < num_cols:
                row.append('')

        # Create header row
        header = rows_data[0]
        self.markdown_lines.append('| ' + ' | '.join(header) + ' |')

        # Create separator row
        separator = '| ' + ' | '.join(['---'] * num_cols) + ' |'
        self.markdown_lines.append(separator)

        # Create data rows
        for row in rows_data[1:]:
            self.markdown_lines.append('| ' + ' | '.join(row) + ' |')

        self.markdown_lines.append('')


def main():
    """Main entry point for standalone usage"""
    import argparse

    parser = argparse.ArgumentParser(
        description='Convert Word documents (.docx) to Markdown files (.md)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s document.docx
  %(prog)s document.docx -o output.md
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

    if args.output and len(args.input_files) > 1:
        print("Error: -o/--output option can only be used with a single input file")
        sys.exit(1)

    for input_file in args.input_files:
        try:
            if args.verbose:
                print(f"Converting: {input_file}")

            converter = DocxToMarkdownConverter(
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
