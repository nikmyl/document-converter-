# Quick Start Guide

## What You Have

A complete Markdown to DOCX converter with two interfaces:
1. **Web Application** - Beautiful drag-and-drop interface
2. **Command Line** - For automation and scripting

## Launching the Web App (Easiest Method)

### Windows
Double-click `start_webapp.bat` or run:
```bash
start_webapp.bat
```

### Mac/Linux
Run:
```bash
./start_webapp.sh
```

### Manual Start
```bash
python app.py
```

Then open your browser to: **http://localhost:5000**

## Using the Web Interface

1. Drag your .md file onto the drop zone (or click "Browse Files")
2. Click "Convert to DOCX"
3. Your file downloads automatically!

That's it! Super simple.

## Command Line Usage

Convert a single file:
```bash
python md_to_docx.py input.md
```

Convert multiple files:
```bash
python md_to_docx.py *.md
```

## What Gets Converted

All markdown formatting is preserved:
- Headings (H1-H6)
- **Bold** and *italic* text
- `Inline code` and code blocks
- Bullet and numbered lists
- Links
- Blockquotes
- And more!

## Need Help?

- **Web App Guide**: See `WEB_APP_GUIDE.md`
- **Command Line Help**: See `README.md`
- **Technical Details**: See `DOCUMENTATION.md`

## File Structure

```
MD to Word/
├── app.py                    # Web application
├── md_to_docx.py            # Command-line script
├── start_webapp.bat         # Windows launcher
├── start_webapp.sh          # Mac/Linux launcher
├── requirements.txt         # Dependencies
├── sample.md                # Test file
│
├── templates/               # Web app HTML
│   └── index.html
│
├── static/                  # Web app assets
│   ├── css/
│   │   └── style.css
│   └── js/
│       └── script.js
│
└── Documentation/
    ├── README.md            # Full usage guide
    ├── WEB_APP_GUIDE.md    # Web app documentation
    ├── DOCUMENTATION.md     # Technical details
    └── QUICKSTART.md       # This file
```

## First Time Setup

Install dependencies (one time only):
```bash
pip install -r requirements.txt
```

This installs:
- python-docx (Word documents)
- markdown (parsing)
- Flask (web server)

## Tips

1. **Test it first**: Try converting `sample.md` to make sure everything works
2. **Web is easier**: Use the web interface for occasional conversions
3. **CLI for automation**: Use command line for batch processing
4. **Check output**: Always open the .docx file to verify formatting
5. **UTF-8 encoding**: Save your markdown files as UTF-8 for best results

## Troubleshooting

**Port already in use?**
- Another app is using port 5000
- Stop the other app or edit `app.py` to change the port

**Cannot access web app?**
- Make sure the server is running
- Try http://127.0.0.1:5000 instead of localhost

**Conversion fails?**
- Check that your .md file has valid markdown syntax
- Ensure file is UTF-8 encoded
- File must be under 16MB

**Dependencies missing?**
```bash
pip install -r requirements.txt
```

## Examples

### Web App
1. Start: `start_webapp.bat` (Windows) or `./start_webapp.sh` (Mac/Linux)
2. Open: http://localhost:5000
3. Drag: Drop your .md file
4. Convert: Click the button
5. Done: File downloads automatically

### Command Line
```bash
# Single file
python md_to_docx.py notes.md

# Custom output name
python md_to_docx.py notes.md -o formatted_notes.docx

# Multiple files
python md_to_docx.py chapter1.md chapter2.md chapter3.md

# All markdown files
python md_to_docx.py *.md
```

## Access from Other Devices

The web app can be accessed from other devices on your network:

1. Find your computer's IP address:
   - Windows: `ipconfig`
   - Mac/Linux: `ifconfig` or `ip addr`

2. On another device, open browser to:
   ```
   http://YOUR_IP_ADDRESS:5000
   ```
   Example: http://192.168.1.100:5000

3. Make sure your firewall allows port 5000

## Production Use

The current setup is for local/development use. For production:

1. Don't use debug mode
2. Use a production WSGI server (gunicorn, waitress)
3. Set up HTTPS with a reverse proxy
4. Add authentication if needed

See `WEB_APP_GUIDE.md` for production deployment details.

---

**Ready to convert?** Start the web app and drag in a markdown file!
