#!/usr/bin/env python3
"""
Markdown to LaTeX Converter
Converts Markdown (.md) files to LaTeX documents (.tex)
"""

import re
import sys
import argparse
from pathlib import Path
from typing import Optional, List


class MarkdownToTexConverter:
    """Convert Markdown files to LaTeX documents"""

    # Supported input extensions
    SUPPORTED_EXTENSIONS = {'.md', '.markdown', '.txt'}

    def __init__(self, input_file: str, output_file: Optional[str] = None):
        """
        Initialize the converter

        Args:
            input_file: Path to the input .md, .markdown, or .txt file
            output_file: Path to the output .tex file (optional)
        """
        self.input_file = Path(input_file)

        if not self.input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")

        if self.input_file.suffix.lower() not in self.SUPPORTED_EXTENSIONS:
            raise ValueError("Input file must be a .md, .markdown, or .txt file")

        if output_file:
            self.output_file = Path(output_file)
        else:
            self.output_file = self.input_file.with_suffix('.tex')

    def convert(self) -> str:
        """
        Convert the Markdown file to LaTeX

        Returns:
            Path to the output file
        """
        # Read the markdown content
        with open(self.input_file, 'r', encoding='utf-8') as f:
            md_content = f.read()

        # Process the markdown and generate LaTeX
        latex_body = self._process_markdown(md_content)

        # Wrap in document template
        latex_content = self._create_document(latex_body)

        # Write output file
        with open(self.output_file, 'w', encoding='utf-8') as f:
            f.write(latex_content)

        return str(self.output_file)

    def _create_document(self, body: str) -> str:
        """Create a complete LaTeX document with preamble"""
        return f"""\\documentclass{{article}}
\\usepackage[utf8]{{inputenc}}
\\usepackage[T1]{{fontenc}}
\\usepackage{{hyperref}}
\\usepackage{{graphicx}}
\\usepackage{{listings}}
\\usepackage{{xcolor}}
\\usepackage{{longtable}}
\\usepackage{{booktabs}}
\\usepackage{{geometry}}

% Page geometry
\\geometry{{margin=1in}}

% Hyperlink setup
\\hypersetup{{
    colorlinks=true,
    linkcolor=blue,
    filecolor=magenta,
    urlcolor=cyan,
}}

% Code listing style
\\lstset{{
    basicstyle=\\ttfamily\\small,
    breaklines=true,
    frame=single,
    backgroundcolor=\\color{{gray!10}},
    xleftmargin=0.5cm,
    xrightmargin=0.5cm,
}}

\\begin{{document}}

{body}

\\end{{document}}
"""

    def _process_markdown(self, content: str) -> str:
        """Process markdown content and return LaTeX body"""
        lines = content.split('\n')
        output_lines = []
        i = 0
        in_code_block = False
        code_block_lines = []
        code_language = ''
        in_list = False
        list_type = None
        in_table = False
        table_rows = []

        while i < len(lines):
            line = lines[i]

            # Handle code blocks
            code_match = re.match(r'^```(\w*)$', line.strip())
            if code_match is not None:
                if not in_code_block:
                    # Close any open list
                    if in_list:
                        output_lines.append(self._close_list(list_type))
                        in_list = False
                    in_code_block = True
                    code_block_lines = []
                    code_language = code_match.group(1)
                else:
                    # End of code block
                    output_lines.append(self._create_code_block(code_block_lines, code_language))
                    in_code_block = False
                    code_block_lines = []
                    code_language = ''
                i += 1
                continue

            if in_code_block:
                code_block_lines.append(line)
                i += 1
                continue

            # Handle tables
            if '|' in line and line.strip().startswith('|'):
                # Close any open list
                if in_list:
                    output_lines.append(self._close_list(list_type))
                    in_list = False
                if not in_table:
                    in_table = True
                    table_rows = []
                table_rows.append(line)
                i += 1
                continue
            elif in_table:
                # End of table
                output_lines.append(self._create_table(table_rows))
                in_table = False
                table_rows = []
                # Don't increment i, process current line

            # Handle headers
            header_match = re.match(r'^(#{1,6})\s+(.+)$', line)
            if header_match:
                # Close any open list
                if in_list:
                    output_lines.append(self._close_list(list_type))
                    in_list = False
                level = len(header_match.group(1))
                text = self._process_inline_formatting(header_match.group(2))
                output_lines.append(self._create_heading(text, level))
                i += 1
                continue

            # Handle horizontal rules
            if re.match(r'^(\*{3,}|-{3,}|_{3,})$', line.strip()):
                # Close any open list
                if in_list:
                    output_lines.append(self._close_list(list_type))
                    in_list = False
                output_lines.append('\n\\noindent\\rule{\\textwidth}{0.4pt}\n')
                i += 1
                continue

            # Handle unordered lists
            list_match = re.match(r'^[\*\-\+]\s+(.+)$', line)
            if list_match:
                if not in_list or list_type != 'itemize':
                    if in_list:
                        output_lines.append(self._close_list(list_type))
                    output_lines.append('\\begin{itemize}')
                    in_list = True
                    list_type = 'itemize'
                text = self._process_inline_formatting(list_match.group(1))
                output_lines.append(f'  \\item {text}')
                i += 1
                continue

            # Handle ordered lists
            ordered_match = re.match(r'^(\d+)\.\s+(.+)$', line)
            if ordered_match:
                if not in_list or list_type != 'enumerate':
                    if in_list:
                        output_lines.append(self._close_list(list_type))
                    output_lines.append('\\begin{enumerate}')
                    in_list = True
                    list_type = 'enumerate'
                text = self._process_inline_formatting(ordered_match.group(2))
                output_lines.append(f'  \\item {text}')
                i += 1
                continue

            # Handle blockquotes
            if line.strip().startswith('>'):
                # Close any open list
                if in_list:
                    output_lines.append(self._close_list(list_type))
                    in_list = False
                quote_lines = []
                while i < len(lines) and lines[i].strip().startswith('>'):
                    quote_text = lines[i].strip()[1:].strip()
                    quote_lines.append(self._process_inline_formatting(quote_text))
                    i += 1
                output_lines.append('\\begin{quote}')
                output_lines.extend(quote_lines)
                output_lines.append('\\end{quote}')
                continue

            # Handle empty lines
            if not line.strip():
                # Close any open list
                if in_list:
                    output_lines.append(self._close_list(list_type))
                    in_list = False
                output_lines.append('')
                i += 1
                continue

            # Handle regular paragraphs
            # Close any open list
            if in_list:
                output_lines.append(self._close_list(list_type))
                in_list = False
            text = self._process_inline_formatting(line)
            output_lines.append(text)
            i += 1

        # Close any remaining structures
        if in_list:
            output_lines.append(self._close_list(list_type))
        if in_table and table_rows:
            output_lines.append(self._create_table(table_rows))

        return '\n'.join(output_lines)

    def _create_heading(self, text: str, level: int) -> str:
        """Create a LaTeX heading command"""
        heading_commands = {
            1: 'section',
            2: 'subsection',
            3: 'subsubsection',
            4: 'paragraph',
            5: 'subparagraph',
            6: 'subparagraph',
        }
        command = heading_commands.get(level, 'paragraph')
        return f'\\{command}{{{text}}}'

    def _process_inline_formatting(self, text: str) -> str:
        """Process inline markdown formatting and return LaTeX markup"""
        # Escape special LaTeX characters first (except those we'll process)
        # Be careful not to escape characters inside markdown formatting

        # Temporarily protect markdown formatting
        text = self._escape_latex_special_chars(text)

        # Inline code: `code` -> \texttt{code}
        text = re.sub(r'`([^`]+)`', r'\\texttt{\1}', text)

        # Links: [text](url) -> \href{url}{text}
        text = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'\\href{\2}{\1}', text)

        # Images: ![alt](url) -> (not embedded, just text)
        text = re.sub(r'!\[([^\]]*)\]\([^)]+\)', r'[Image: \1]', text)

        # Bold and italic: ***text*** or ___text___
        text = re.sub(r'\*\*\*([^*]+)\*\*\*', r'\\textbf{\\textit{\1}}', text)
        text = re.sub(r'___([^_]+)___', r'\\textbf{\\textit{\1}}', text)

        # Bold: **text** or __text__
        text = re.sub(r'\*\*([^*]+)\*\*', r'\\textbf{\1}', text)
        text = re.sub(r'__([^_]+)__', r'\\textbf{\1}', text)

        # Italic: *text* or _text_
        text = re.sub(r'(?<!\*)\*([^*]+)\*(?!\*)', r'\\textit{\1}', text)
        text = re.sub(r'(?<!_)_([^_]+)_(?!_)', r'\\textit{\1}', text)

        # Strikethrough: ~~text~~ -> \sout{text} (requires ulem package, skip for now)
        text = re.sub(r'~~([^~]+)~~', r'\1', text)

        return text

    def _escape_latex_special_chars(self, text: str) -> str:
        """Escape special LaTeX characters"""
        # Characters that need escaping in LaTeX
        # Note: We handle \ first to avoid double-escaping
        # We don't escape * and _ here as they're used for markdown formatting

        # First protect existing LaTeX commands (backslash followed by letters)
        # by temporarily replacing them
        protected = []
        def protect_commands(match):
            protected.append(match.group(0))
            return f'\x00PROTECTED{len(protected)-1}\x00'

        text = re.sub(r'\\[a-zA-Z]+', protect_commands, text)

        # Escape special characters
        text = text.replace('\\', '\\textbackslash{}')
        text = text.replace('&', '\\&')
        text = text.replace('%', '\\%')
        text = text.replace('$', '\\$')
        text = text.replace('#', '\\#')
        text = text.replace('{', '\\{')
        text = text.replace('}', '\\}')
        text = text.replace('~', '\\textasciitilde{}')
        text = text.replace('^', '\\textasciicircum{}')

        # Restore protected commands
        for i, cmd in enumerate(protected):
            text = text.replace(f'\x00PROTECTED{i}\x00', cmd)

        return text

    def _create_code_block(self, lines: List[str], language: str = '') -> str:
        """Create a LaTeX code block"""
        code = '\n'.join(lines)
        if language:
            return f'\\begin{{lstlisting}}[language={language}]\n{code}\n\\end{{lstlisting}}'
        else:
            return f'\\begin{{lstlisting}}\n{code}\n\\end{{lstlisting}}'

    def _close_list(self, list_type: str) -> str:
        """Close a list environment"""
        return f'\\end{{{list_type}}}'

    def _create_table(self, table_rows: List[str]) -> str:
        """Create a LaTeX table from markdown table rows"""
        if not table_rows:
            return ''

        # Parse table rows
        parsed_rows = []
        for row in table_rows:
            cells = [cell.strip() for cell in row.split('|')]
            if cells and cells[0] == '':
                cells = cells[1:]
            if cells and cells[-1] == '':
                cells = cells[:-1]

            # Skip separator row (contains only - and :)
            if cells and all(set(cell.replace('-', '').replace(':', '').replace(' ', '')) == set() for cell in cells):
                continue

            if cells:
                # Process inline formatting in cells
                cells = [self._process_inline_formatting(cell) for cell in cells]
                parsed_rows.append(cells)

        if not parsed_rows:
            return ''

        # Determine column count
        num_cols = max(len(row) for row in parsed_rows)

        # Create column specification
        col_spec = '|' + 'l|' * num_cols

        # Build table
        output = []
        output.append(f'\\begin{{tabular}}{{{col_spec}}}')
        output.append('\\hline')

        for i, row in enumerate(parsed_rows):
            # Pad row to num_cols
            while len(row) < num_cols:
                row.append('')
            output.append(' & '.join(row) + ' \\\\')
            output.append('\\hline')

        output.append('\\end{tabular}')
        output.append('')  # Add blank line after table

        return '\n'.join(output)


def main():
    """Main entry point for the script"""
    parser = argparse.ArgumentParser(
        description='Convert Markdown files to LaTeX documents (.tex)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s input.md
  %(prog)s input.md -o output.tex
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

            converter = MarkdownToTexConverter(
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
