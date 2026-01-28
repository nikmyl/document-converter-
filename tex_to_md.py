#!/usr/bin/env python3
"""
LaTeX to Markdown Converter
Converts LaTeX (.tex) files to Markdown documents (.md)
"""

import re
import sys
import argparse
from pathlib import Path
from typing import Optional, List, Tuple


class TexToMarkdownConverter:
    """Convert LaTeX files to Markdown documents"""

    # Supported input extensions
    SUPPORTED_EXTENSIONS = {'.tex', '.latex'}

    def __init__(self, input_file: str, output_file: Optional[str] = None):
        """
        Initialize the converter

        Args:
            input_file: Path to the input .tex or .latex file
            output_file: Path to the output .md file (optional)
        """
        self.input_file = Path(input_file)

        if not self.input_file.exists():
            raise FileNotFoundError(f"Input file not found: {input_file}")

        if self.input_file.suffix.lower() not in self.SUPPORTED_EXTENSIONS:
            raise ValueError("Input file must be a .tex or .latex file")

        if output_file:
            self.output_file = Path(output_file)
        else:
            self.output_file = self.input_file.with_suffix('.md')

    def convert(self) -> str:
        """
        Convert the LaTeX file to Markdown

        Returns:
            Path to the output file
        """
        # Read the LaTeX content
        with open(self.input_file, 'r', encoding='utf-8') as f:
            tex_content = f.read()

        # Process the LaTeX and generate Markdown
        md_content = self._process_latex(tex_content)

        # Write output file
        with open(self.output_file, 'w', encoding='utf-8') as f:
            f.write(md_content)

        return str(self.output_file)

    def _process_latex(self, content: str) -> str:
        """Process LaTeX content and return Markdown"""
        # Extract document body (between \begin{document} and \end{document})
        body_match = re.search(
            r'\\begin\{document\}(.*?)\\end\{document\}',
            content,
            re.DOTALL
        )

        if body_match:
            content = body_match.group(1)

        # Remove common preamble commands that might remain
        content = re.sub(r'\\maketitle', '', content)
        content = re.sub(r'\\tableofcontents', '', content)

        # Process LaTeX elements
        content = self._process_headings(content)
        content = self._process_formatting(content)
        content = self._process_links(content)
        content = self._process_lists(content)
        content = self._process_code_blocks(content)
        content = self._process_blockquotes(content)
        content = self._process_tables(content)
        content = self._process_special_chars(content)
        content = self._cleanup(content)

        return content.strip()

    def _process_headings(self, content: str) -> str:
        """Convert LaTeX headings to Markdown"""
        # \section{Title} -> # Title
        content = re.sub(r'\\section\*?\{([^}]+)\}', r'# \1', content)
        # \subsection{Title} -> ## Title
        content = re.sub(r'\\subsection\*?\{([^}]+)\}', r'## \1', content)
        # \subsubsection{Title} -> ### Title
        content = re.sub(r'\\subsubsection\*?\{([^}]+)\}', r'### \1', content)
        # \paragraph{Title} -> #### Title
        content = re.sub(r'\\paragraph\*?\{([^}]+)\}', r'#### \1', content)
        # \subparagraph{Title} -> ##### Title
        content = re.sub(r'\\subparagraph\*?\{([^}]+)\}', r'##### \1', content)
        # \chapter{Title} -> # Title (for book classes)
        content = re.sub(r'\\chapter\*?\{([^}]+)\}', r'# \1', content)

        return content

    def _process_formatting(self, content: str) -> str:
        """Convert LaTeX text formatting to Markdown"""
        # Handle nested bold+italic first
        content = re.sub(r'\\textbf\{\\textit\{([^}]+)\}\}', r'***\1***', content)
        content = re.sub(r'\\textit\{\\textbf\{([^}]+)\}\}', r'***\1***', content)
        content = re.sub(r'\\emph\{\\textbf\{([^}]+)\}\}', r'***\1***', content)
        content = re.sub(r'\\textbf\{\\emph\{([^}]+)\}\}', r'***\1***', content)

        # \textbf{text} -> **text**
        content = re.sub(r'\\textbf\{([^}]+)\}', r'**\1**', content)
        # \textit{text} -> *text*
        content = re.sub(r'\\textit\{([^}]+)\}', r'*\1*', content)
        # \emph{text} -> *text*
        content = re.sub(r'\\emph\{([^}]+)\}', r'*\1*', content)
        # \underline{text} -> text (no direct markdown equivalent)
        content = re.sub(r'\\underline\{([^}]+)\}', r'\1', content)
        # \texttt{text} -> `text`
        content = re.sub(r'\\texttt\{([^}]+)\}', r'`\1`', content)
        # \verb|text| -> `text`
        content = re.sub(r'\\verb\|([^|]+)\|', r'`\1`', content)
        content = re.sub(r'\\verb\+([^+]+)\+', r'`\1`', content)
        # \sout{text} or \st{text} -> ~~text~~ (strikethrough)
        content = re.sub(r'\\sout\{([^}]+)\}', r'~~\1~~', content)
        content = re.sub(r'\\st\{([^}]+)\}', r'~~\1~~', content)

        return content

    def _process_links(self, content: str) -> str:
        """Convert LaTeX links to Markdown"""
        # \href{url}{text} -> [text](url)
        content = re.sub(r'\\href\{([^}]+)\}\{([^}]+)\}', r'[\2](\1)', content)
        # \url{url} -> <url>
        content = re.sub(r'\\url\{([^}]+)\}', r'<\1>', content)

        return content

    def _process_lists(self, content: str) -> str:
        """Convert LaTeX lists to Markdown"""
        # Process itemize (unordered lists)
        def convert_itemize(match):
            items = match.group(1)
            # Split by \item and process
            parts = re.split(r'\\item\s*', items)
            result = []
            for part in parts:
                part = part.strip()
                if part:
                    # Handle multi-line items
                    lines = part.split('\n')
                    first_line = lines[0].strip()
                    if first_line:
                        result.append(f'- {first_line}')
                    for line in lines[1:]:
                        line = line.strip()
                        if line:
                            result.append(f'  {line}')
            return '\n'.join(result)

        content = re.sub(
            r'\\begin\{itemize\}(.*?)\\end\{itemize\}',
            convert_itemize,
            content,
            flags=re.DOTALL
        )

        # Process enumerate (ordered lists)
        def convert_enumerate(match):
            items = match.group(1)
            parts = re.split(r'\\item\s*', items)
            result = []
            num = 1
            for part in parts:
                part = part.strip()
                if part:
                    lines = part.split('\n')
                    first_line = lines[0].strip()
                    if first_line:
                        result.append(f'{num}. {first_line}')
                        num += 1
                    for line in lines[1:]:
                        line = line.strip()
                        if line:
                            result.append(f'   {line}')
            return '\n'.join(result)

        content = re.sub(
            r'\\begin\{enumerate\}(.*?)\\end\{enumerate\}',
            convert_enumerate,
            content,
            flags=re.DOTALL
        )

        # Process description lists
        def convert_description(match):
            items = match.group(1)
            # Match \item[term] description
            result = []
            for item_match in re.finditer(r'\\item\[([^\]]+)\]\s*([^\\]+)', items):
                term = item_match.group(1).strip()
                desc = item_match.group(2).strip()
                result.append(f'**{term}**: {desc}')
            return '\n'.join(result)

        content = re.sub(
            r'\\begin\{description\}(.*?)\\end\{description\}',
            convert_description,
            content,
            flags=re.DOTALL
        )

        return content

    def _process_code_blocks(self, content: str) -> str:
        """Convert LaTeX code environments to Markdown code blocks"""
        # \begin{verbatim}...\end{verbatim} -> ```...```
        def convert_verbatim(match):
            code = match.group(1)
            return f'```\n{code.strip()}\n```'

        content = re.sub(
            r'\\begin\{verbatim\}(.*?)\\end\{verbatim\}',
            convert_verbatim,
            content,
            flags=re.DOTALL
        )

        # \begin{lstlisting}[language=X]...\end{lstlisting} -> ```X...```
        def convert_lstlisting(match):
            options = match.group(1) or ''
            code = match.group(2)
            # Extract language if specified
            lang_match = re.search(r'language=(\w+)', options)
            lang = lang_match.group(1).lower() if lang_match else ''
            return f'```{lang}\n{code.strip()}\n```'

        content = re.sub(
            r'\\begin\{lstlisting\}(?:\[([^\]]*)\])?(.*?)\\end\{lstlisting\}',
            convert_lstlisting,
            content,
            flags=re.DOTALL
        )

        # \begin{minted}{language}...\end{minted} -> ```language...```
        def convert_minted(match):
            lang = match.group(1).lower()
            code = match.group(2)
            return f'```{lang}\n{code.strip()}\n```'

        content = re.sub(
            r'\\begin\{minted\}\{(\w+)\}(.*?)\\end\{minted\}',
            convert_minted,
            content,
            flags=re.DOTALL
        )

        return content

    def _process_blockquotes(self, content: str) -> str:
        """Convert LaTeX quote environments to Markdown blockquotes"""
        def convert_quote(match):
            quote_text = match.group(1).strip()
            lines = quote_text.split('\n')
            return '\n'.join(f'> {line.strip()}' for line in lines if line.strip())

        content = re.sub(
            r'\\begin\{quote\}(.*?)\\end\{quote\}',
            convert_quote,
            content,
            flags=re.DOTALL
        )

        content = re.sub(
            r'\\begin\{quotation\}(.*?)\\end\{quotation\}',
            convert_quote,
            content,
            flags=re.DOTALL
        )

        return content

    def _process_tables(self, content: str) -> str:
        """Convert LaTeX tables to Markdown tables"""
        def convert_tabular(match):
            col_spec = match.group(1)
            table_content = match.group(2)

            # Count columns from col_spec
            num_cols = len(re.findall(r'[lcr]', col_spec))
            if num_cols == 0:
                num_cols = len(re.findall(r'[|]', col_spec)) - 1
                if num_cols <= 0:
                    num_cols = 1

            # Parse rows
            rows = []
            # Split by \\ (row separator)
            row_texts = re.split(r'\\\\', table_content)

            for row_text in row_texts:
                row_text = row_text.strip()
                # Skip empty rows and \hline
                if not row_text or row_text == '\\hline' or re.match(r'^\\[a-z]+line$', row_text):
                    continue
                # Remove \hline from start/end
                row_text = re.sub(r'\\hline\s*', '', row_text)
                row_text = row_text.strip()
                if not row_text:
                    continue

                # Split by & (column separator)
                cells = [cell.strip() for cell in row_text.split('&')]
                rows.append(cells)

            if not rows:
                return ''

            # Build markdown table
            result = []
            # Normalize row lengths
            max_cols = max(len(row) for row in rows)
            for row in rows:
                while len(row) < max_cols:
                    row.append('')

            # Header row
            result.append('| ' + ' | '.join(rows[0]) + ' |')
            # Separator row
            result.append('| ' + ' | '.join(['---'] * max_cols) + ' |')
            # Data rows
            for row in rows[1:]:
                result.append('| ' + ' | '.join(row) + ' |')

            return '\n'.join(result)

        # Match tabular environment
        content = re.sub(
            r'\\begin\{tabular\}\{([^}]+)\}(.*?)\\end\{tabular\}',
            convert_tabular,
            content,
            flags=re.DOTALL
        )

        # Also handle longtable
        content = re.sub(
            r'\\begin\{longtable\}\{([^}]+)\}(.*?)\\end\{longtable\}',
            convert_tabular,
            content,
            flags=re.DOTALL
        )

        # Remove table environment wrapper if present
        content = re.sub(r'\\begin\{table\}\[?[^\]]*\]?', '', content)
        content = re.sub(r'\\end\{table\}', '', content)
        content = re.sub(r'\\centering', '', content)
        content = re.sub(r'\\caption\{[^}]*\}', '', content)
        content = re.sub(r'\\label\{[^}]*\}', '', content)

        return content

    def _process_special_chars(self, content: str) -> str:
        """Convert LaTeX special character escapes back to regular characters"""
        # Escaped characters
        content = content.replace('\\&', '&')
        content = content.replace('\\%', '%')
        content = content.replace('\\$', '$')
        content = content.replace('\\#', '#')
        content = content.replace('\\{', '{')
        content = content.replace('\\}', '}')
        content = content.replace('\\textbackslash{}', '\\')
        content = content.replace('\\textbackslash', '\\')
        content = content.replace('\\textasciitilde{}', '~')
        content = content.replace('\\textasciitilde', '~')
        content = content.replace('\\textasciicircum{}', '^')
        content = content.replace('\\textasciicircum', '^')

        # Special LaTeX characters
        content = content.replace('~', ' ')  # Non-breaking space
        content = content.replace('\\\\', '\n')  # Line break
        content = content.replace('\\newline', '\n')
        content = content.replace('\\par', '\n\n')

        # Quotes
        content = content.replace('``', '"')
        content = content.replace("''", '"')
        content = content.replace('`', "'")

        # Dashes
        content = content.replace('---', '\u2014')  # em dash
        content = content.replace('--', '\u2013')   # en dash

        # Horizontal rule
        content = re.sub(r'\\noindent\\rule\{[^}]*\}\{[^}]*\}', '\n---\n', content)
        content = re.sub(r'\\rule\{[^}]*\}\{[^}]*\}', '\n---\n', content)
        content = re.sub(r'\\hrule', '\n---\n', content)

        return content

    def _cleanup(self, content: str) -> str:
        """Clean up the converted content"""
        # Remove remaining LaTeX commands we don't handle
        content = re.sub(r'\\[a-zA-Z]+\{[^}]*\}', '', content)
        content = re.sub(r'\\[a-zA-Z]+\[[^\]]*\]', '', content)
        content = re.sub(r'\\[a-zA-Z]+', '', content)

        # Remove multiple blank lines
        content = re.sub(r'\n{3,}', '\n\n', content)

        # Remove leading/trailing whitespace from lines
        lines = content.split('\n')
        lines = [line.rstrip() for line in lines]
        content = '\n'.join(lines)

        return content


def main():
    """Main entry point for the script"""
    parser = argparse.ArgumentParser(
        description='Convert LaTeX files to Markdown documents (.md)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s input.tex
  %(prog)s input.tex -o output.md
  %(prog)s *.tex
        """
    )

    parser.add_argument(
        'input_files',
        nargs='+',
        help='Input LaTeX file(s) to convert'
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

            converter = TexToMarkdownConverter(
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
