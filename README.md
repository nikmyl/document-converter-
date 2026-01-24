# Markdown to DOCX Converter

A Python tool that converts Markdown (.md) files to properly formatted Microsoft Word documents (.docx).

**Two Ways to Use:**
1. **Web Interface** - Drag-and-drop web application (recommended for most users)
2. **Command Line** - Terminal-based conversion for automation and batch processing

## Features

- Converts markdown files to Word documents while preserving formatting
- Supports headings, bold, italic, inline code, code blocks, lists, links, blockquotes, and **tables**
- **Web interface** with drag-and-drop functionality
- **Command-line interface** for automation
- Batch conversion support (convert multiple files at once)
- Cross-platform compatibility (Windows, macOS, Linux)

## Requirements

- Python 3.7 or higher
- pip (Python package installer)

## Installation

1. **Clone or download this repository** to your local machine

2. **Install the required dependencies**:

   ```bash
   pip install -r requirements.txt
   ```

   This will install:
   - `python-docx` - For creating Word documents
   - `markdown` - For markdown processing
   - `Flask` - For the web application

## Quick Start

### Option 1: Web Interface (Recommended)

**Start the web application:**

Windows:
```bash
start_webapp.bat
```

Mac/Linux:
```bash
./start_webapp.sh
```

Or manually:
```bash
python app.py
```

Then open your browser and go to **http://localhost:5000**

**Using the web interface:**
1. Drag and drop your .md file onto the drop zone, or click "Browse Files"
2. Click "Convert to DOCX"
3. The converted file will download automatically

See `WEB_APP_GUIDE.md` for detailed web interface documentation.

### Option 2: Command Line

Convert a single markdown file:

```bash
python md_to_docx.py input.md
```

This will create `input.docx` in the same directory.

## Usage

### Basic Usage

```bash
python md_to_docx.py <input_file.md>
```

Example:
```bash
python md_to_docx.py README.md
```

### Specify Output File

```bash
python md_to_docx.py input.md -o output.docx
```

Example:
```bash
python md_to_docx.py notes.md -o formatted_notes.docx
```

### Convert Multiple Files

```bash
python md_to_docx.py file1.md file2.md file3.md
```

Or use wildcards:
```bash
python md_to_docx.py *.md
```

### Verbose Mode

Get detailed output during conversion:

```bash
python md_to_docx.py input.md -v
```

## Command-Line Options

| Option | Description |
|--------|-------------|
| `input_files` | One or more markdown files to convert (required) |
| `-o, --output` | Specify output file path (only works with single input file) |
| `-v, --verbose` | Enable verbose output for debugging |
| `-h, --help` | Show help message |

## Supported Markdown Features

The converter supports the following markdown elements:

### Text Formatting
- **Bold text**: `**bold**` or `__bold__`
- *Italic text*: `*italic*` or `_italic_`
- `Inline code`: `` `code` ``

### Headings
```markdown
# Heading 1
## Heading 2
### Heading 3
#### Heading 4
##### Heading 5
###### Heading 6
```

### Lists

Unordered lists:
```markdown
* Item 1
* Item 2
* Item 3
```

Ordered lists:
```markdown
1. First item
2. Second item
3. Third item
```

### Code Blocks

Use triple backticks:

````markdown
```python
def hello():
    print("Hello, World!")
```
````

### Links

```markdown
[Link text](https://www.example.com)
```

### Blockquotes

```markdown
> This is a blockquote
> It can span multiple lines
```

### Horizontal Rules

```markdown
---
```

### Tables

```markdown
| Column 1 | Column 2 | Column 3 |
|----------|----------|----------|
| Data 1   | Data 2   | Data 3   |
| Data 4   | Data 5   | Data 6   |
```

Tables are fully supported with:
- Header row formatting (bold)
- Proper cell alignment
- Empty cells supported
- Professional table styling

## Examples

### Example 1: Convert a Single File

```bash
python md_to_docx.py report.md
```

Output: `report.docx` created in the same directory

### Example 2: Convert with Custom Output Name

```bash
python md_to_docx.py draft.md -o final_report.docx
```

Output: `final_report.docx` created

### Example 3: Convert All Markdown Files in Directory

```bash
python md_to_docx.py *.md -v
```

Output: Creates .docx version of each .md file with verbose progress

### Example 4: Test with Sample File

A sample markdown file is included for testing:

```bash
python md_to_docx.py sample.md
```

Open `sample.docx` to see how various markdown elements are formatted.

## Output Format

The generated Word documents include:

- Properly styled headings (Heading 1-6)
- Formatted text (bold, italic, inline code)
- Bullet and numbered lists
- Code blocks with monospace font and indentation
- Blockquotes with Quote style
- Links formatted as blue underlined text
- Professional spacing and formatting

## Troubleshooting

### Issue: "python-docx is required"

**Solution**: Install dependencies
```bash
pip install -r requirements.txt
```

### Issue: "File not found"

**Solution**: Make sure the file path is correct. Use quotes for paths with spaces:
```bash
python md_to_docx.py "my documents/file.md"
```

### Issue: Python command not found

**Solution**: Make sure Python is installed and in your PATH. Try:
```bash
python3 md_to_docx.py input.md
```

or

```bash
py md_to_docx.py input.md
```

### Issue: Output file is empty or improperly formatted

**Solution**: Check that your markdown file:
- Is saved in UTF-8 encoding
- Has valid markdown syntax
- Is not empty

## Limitations

Current limitations:
- Images are not embedded (only link text is shown)
- No syntax highlighting for code blocks
- HTML within markdown is not processed
- Complex nested formatting may not render perfectly

See `DOCUMENTATION.md` for detailed technical information and future enhancement plans.

## Files Included

### Core Files
- `md_to_docx.py` - Main conversion script
- `app.py` - Flask web application
- `requirements.txt` - Python dependencies
- `sample.md` - Sample markdown file for testing

### Web Application Files
- `templates/index.html` - Web interface HTML
- `static/css/style.css` - Web interface styling
- `static/js/script.js` - Client-side JavaScript
- `start_webapp.bat` - Windows startup script
- `start_webapp.sh` - Mac/Linux startup script

### Documentation
- `README.md` - This file (usage instructions)
- `WEB_APP_GUIDE.md` - Web application guide
- `DOCUMENTATION.md` - Technical documentation

## Tips for Best Results

1. **Use standard markdown syntax** - The converter works best with standard markdown
2. **Test with sample.md first** - Verify the installation works correctly
3. **Check encoding** - Save markdown files in UTF-8 encoding
4. **Preview output** - Always open the generated .docx file to verify formatting
5. **Use verbose mode** - Add `-v` flag when troubleshooting issues

## Integration with Other Tools

### Use with Git

Convert your markdown documentation to Word for sharing:
```bash
python md_to_docx.py README.md -o ProjectDocumentation.docx
```

### Batch Processing

Create a batch script to convert all markdown files in a project:

**Windows (batch file)**:
```batch
@echo off
for %%f in (*.md) do python md_to_docx.py "%%f"
```

**Linux/Mac (shell script)**:
```bash
#!/bin/bash
for file in *.md; do
    python md_to_docx.py "$file"
done
```

### Use in Python Scripts

You can also import and use the converter in your own Python scripts:

```python
from md_to_docx import MarkdownToDocxConverter

converter = MarkdownToDocxConverter("input.md", "output.docx")
output_path = converter.convert()
print(f"Converted to: {output_path}")
```

## Getting Help

- For usage help: `python md_to_docx.py --help`
- For technical details: See `DOCUMENTATION.md`
- For examples: Check `sample.md` and its output

## Contributing

Feel free to modify and extend this script for your needs. Some ideas for enhancements:
- Add table support
- Embed images in the document
- Add syntax highlighting for code blocks
- Create custom styling themes
- Add progress bars for batch conversions

## License

This is an open-source tool. Feel free to use, modify, and distribute.

## Version

Current version: 1.0

---

**Happy converting!** If you encounter any issues or have suggestions, please check the DOCUMENTATION.md file for more details.
