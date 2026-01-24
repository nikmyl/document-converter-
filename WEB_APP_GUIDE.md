# Markdown to DOCX Web Application Guide

## Overview

The web application provides an easy-to-use interface for converting Markdown files to Word documents with drag-and-drop functionality.

## Starting the Web Application

### Option 1: Using Python directly

```bash
python app.py
```

### Option 2: Using Flask command

```bash
flask run
```

The application will start on **http://localhost:5000**

You'll see output like:
```
============================================================
Markdown to DOCX Converter - Web Interface
============================================================

Starting server...
Access the application at: http://localhost:5000

Press CTRL+C to stop the server
============================================================
```

## Accessing the Application

1. Open your web browser
2. Navigate to: **http://localhost:5000**
3. You'll see the Markdown to DOCX Converter interface

## Using the Web Interface

### Method 1: Drag and Drop

1. Drag a Markdown file (.md, .markdown, or .txt) from your file explorer
2. Drop it onto the "Drag & Drop" zone
3. The file information will appear
4. Click the "Convert to DOCX" button
5. The converted file will download automatically

### Method 2: Browse Files

1. Click the "Browse Files" button
2. Select a Markdown file from your computer
3. The file information will appear
4. Click the "Convert to DOCX" button
5. The converted file will download automatically

## Features

### Supported File Types
- `.md` - Markdown files
- `.markdown` - Markdown files
- `.txt` - Text files with markdown formatting

### File Size Limit
- Maximum file size: **16 MB**

### Security
- Files are processed locally on your server
- No data is sent to external services
- Temporary files are automatically cleaned up

### Conversion Features
- Preserves all markdown formatting
- Headings (H1-H6)
- Bold and italic text
- Inline code and code blocks
- Bullet and numbered lists
- Links
- Blockquotes
- Horizontal rules

## Stopping the Server

To stop the web application:
- Press **CTRL+C** in the terminal where the server is running

## Configuration

### Changing the Port

Edit `app.py` and modify the last line:

```python
app.run(debug=True, host='0.0.0.0', port=5000)
```

Change `port=5000` to your desired port number.

### Changing the File Size Limit

Edit `app.py` and modify this line:

```python
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
```

Change the value to your desired size (in bytes).

### Running on a Specific IP Address

By default, the app runs on `0.0.0.0` which makes it accessible from:
- `http://localhost:5000`
- `http://127.0.0.1:5000`
- `http://YOUR_LOCAL_IP:5000` (from other devices on your network)

To restrict to localhost only, change `host='0.0.0.0'` to `host='127.0.0.1'`.

## Accessing from Other Devices on Your Network

If you want to access the converter from other devices on your local network:

1. Find your computer's local IP address:
   - **Windows**: Run `ipconfig` in Command Prompt, look for "IPv4 Address"
   - **Mac/Linux**: Run `ifconfig` or `ip addr`

2. On the other device, open a browser and navigate to:
   ```
   http://YOUR_LOCAL_IP:5000
   ```
   For example: `http://192.168.1.100:5000`

3. Make sure your firewall allows connections on port 5000

## Troubleshooting

### Issue: Port Already in Use

**Error**: `Address already in use`

**Solution**:
- Either stop the other application using port 5000
- Or change the port in `app.py` to a different number (e.g., 5001, 8000, etc.)

### Issue: Cannot Access from Browser

**Solution**:
1. Check that the server is running (you should see the startup message)
2. Try `http://127.0.0.1:5000` instead of `localhost`
3. Make sure no firewall is blocking the port

### Issue: File Upload Fails

**Solution**:
1. Check that the file is a valid markdown file
2. Ensure the file is under 16MB
3. Check the terminal for error messages

### Issue: Conversion Fails

**Solution**:
1. Check that the markdown file has valid syntax
2. Ensure the file is UTF-8 encoded
3. Look at the error message displayed in the browser
4. Check the terminal for detailed error logs

### Issue: Download Doesn't Start

**Solution**:
1. Check your browser's download settings
2. Ensure pop-ups are not blocked
3. Try a different browser

## Production Deployment

**Important**: The current configuration is for development only. For production deployment:

1. **Never use debug mode in production**
   - Change `debug=True` to `debug=False` in `app.py`

2. **Use a production WSGI server**
   - Install: `pip install gunicorn` (Linux/Mac) or `pip install waitress` (Windows)
   - Run with Gunicorn: `gunicorn -w 4 -b 0.0.0.0:5000 app:app`
   - Run with Waitress: `waitress-serve --port=5000 app:app`

3. **Add proper security measures**
   - Implement rate limiting
   - Add authentication if needed
   - Use HTTPS with a reverse proxy (nginx, Apache)
   - Set proper CORS policies

4. **Configure environment variables**
   ```bash
   export FLASK_APP=app.py
   export FLASK_ENV=production
   ```

## File Structure

```
MD to Word/
├── app.py                 # Flask application
├── md_to_docx.py         # Conversion logic
├── requirements.txt       # Python dependencies
├── templates/
│   └── index.html        # Web interface HTML
├── static/
│   ├── css/
│   │   └── style.css     # Styling
│   └── js/
│       └── script.js     # Client-side logic
└── WEB_APP_GUIDE.md      # This file
```

## API Endpoints

The application provides these endpoints:

### GET /
Returns the main HTML page

### POST /convert
Accepts a markdown file and returns the converted DOCX file

**Request**: `multipart/form-data` with file field named `file`

**Response**: DOCX file as download

**Error codes**:
- 400: Bad request (no file, invalid file type)
- 413: File too large
- 500: Conversion error

### GET /health
Health check endpoint

**Response**: `{"status": "ok"}`

## Tips for Best Experience

1. **Use modern browsers**: Chrome, Firefox, Safari, or Edge (latest versions)
2. **Check file encoding**: Ensure markdown files are UTF-8 encoded
3. **Test with sample**: Use the included `sample.md` file to test the application
4. **Keep files organized**: Downloaded files go to your browser's download folder
5. **Batch processing**: You can convert multiple files one after another

## Integration with Other Tools

### Using curl

```bash
curl -F "file=@sample.md" http://localhost:5000/convert -o output.docx
```

### Using Python requests

```python
import requests

with open('sample.md', 'rb') as f:
    files = {'file': f}
    response = requests.post('http://localhost:5000/convert', files=files)

    with open('output.docx', 'wb') as out:
        out.write(response.content)
```

### Using JavaScript/Node.js

```javascript
const FormData = require('form-data');
const fs = require('fs');
const axios = require('axios');

const form = new FormData();
form.append('file', fs.createReadStream('sample.md'));

axios.post('http://localhost:5000/convert', form, {
    headers: form.getHeaders(),
    responseType: 'arraybuffer'
})
.then(response => {
    fs.writeFileSync('output.docx', response.data);
});
```

## Customizing the Interface

### Change Colors

Edit `static/css/style.css` and modify the color values:

```css
/* Primary color gradient */
background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);

/* Change to your preferred colors */
background: linear-gradient(135deg, #YOUR_COLOR_1 0%, #YOUR_COLOR_2 100%);
```

### Change Text

Edit `templates/index.html` and modify the text content:

```html
<h1>Markdown to DOCX Converter</h1>
<p class="subtitle">Convert your Markdown files to Word documents instantly</p>
```

### Add Logo

Add your logo to the header in `templates/index.html`:

```html
<header>
    <img src="{{ url_for('static', filename='logo.png') }}" alt="Logo">
    <h1>Markdown to DOCX Converter</h1>
</header>
```

And place your logo file in the `static/` directory.

## Performance Considerations

- **Concurrent users**: The development server handles one request at a time
- **Large files**: Files over 5MB may take a few seconds to convert
- **Memory usage**: Each conversion uses temporary disk space
- **Cleanup**: Temporary files are automatically removed

For better performance with multiple users, use a production WSGI server with multiple workers.

## Support

For issues with:
- **Command-line script**: See `README.md`
- **Technical details**: See `DOCUMENTATION.md`
- **Web interface**: This guide

---

**Enjoy converting your Markdown files!**
