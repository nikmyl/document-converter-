// Get DOM elements
const dropZone = document.getElementById('drop-zone');
const dropZoneTitle = document.getElementById('drop-zone-title');
const supportedFormats = document.getElementById('supported-formats');
const fileInput = document.getElementById('file-input');
const fileInputMultiple = document.getElementById('file-input-multiple');
const fileInfoContainer = document.getElementById('file-info');
const batchFileInfo = document.getElementById('batch-file-info');
const fileNameDisplay = document.getElementById('file-name');
const fileSizeDisplay = document.getElementById('file-size');
const convertButtonsContainer = document.getElementById('convert-buttons-container');
const batchConvertButton = document.getElementById('batch-convert-button');
const progressContainer = document.getElementById('progress');
const progressText = document.getElementById('progress-text');
const successMessage = document.getElementById('success');
const successMessageText = document.getElementById('success-message');
const errorMessage = document.getElementById('error');
const errorMessageText = document.getElementById('error-message');
const resetButton = document.getElementById('reset-button');
const errorResetButton = document.getElementById('error-reset-button');
const singleModeBtn = document.getElementById('single-mode-btn');
const batchModeBtn = document.getElementById('batch-mode-btn');
const batchFileList = document.getElementById('batch-file-list');
const batchTitle = document.getElementById('batch-title');
const batchCount = document.getElementById('batch-count');
const mdCount = document.getElementById('md-count');
const docxCount = document.getElementById('docx-count');
const pdfCount = document.getElementById('pdf-count');

// State
let selectedFile = null;
let selectedFiles = [];
let selectedTargetFormat = null;
let isBatchMode = false;
let currentAbortController = null;

// Configuration
const CONFIG = {
    MAX_SINGLE_FILE_SIZE: 16 * 1024 * 1024,  // 16MB
    MAX_BATCH_SIZE: 100 * 1024 * 1024,        // 100MB
    REQUEST_TIMEOUT: 300000,                   // 5 minutes
    RETRY_ATTEMPTS: 2,
    RETRY_DELAY: 1000                          // 1 second
};

// File extension sets
const markdownExtensions = ['md', 'markdown', 'txt'];
const docxExtensions = ['docx'];
const pdfExtensions = ['pdf'];
const zipExtensions = ['zip'];

// Mode toggle handlers
singleModeBtn.addEventListener('click', () => setMode(false));
batchModeBtn.addEventListener('click', () => setMode(true));

function setMode(batch) {
    isBatchMode = batch;

    // Update button states
    singleModeBtn.classList.toggle('active', !batch);
    batchModeBtn.classList.toggle('active', batch);

    // Update aria-pressed states
    singleModeBtn.setAttribute('aria-pressed', !batch);
    batchModeBtn.setAttribute('aria-pressed', batch);

    // Update drop zone text
    if (batch) {
        dropZoneTitle.textContent = 'Drag & Drop files or a .zip here';
        supportedFormats.textContent = 'Supported: .md, .markdown, .txt, .docx, .pdf, .zip (max 100MB total)';
    } else {
        dropZoneTitle.textContent = 'Drag & Drop your file here';
        supportedFormats.textContent = 'Supported: .md, .markdown, .txt, .docx, .pdf (max 16MB)';
    }

    // Reset the form when switching modes
    resetForm();
}

// Prevent default drag behaviors
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, preventDefaults, false);
    document.body.addEventListener(eventName, preventDefaults, false);
});

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

// Highlight drop zone when item is dragged over it
['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, highlight, false);
});

['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, unhighlight, false);
});

function highlight() {
    dropZone.classList.add('drag-over');
}

function unhighlight() {
    dropZone.classList.remove('drag-over');
}

// Handle dropped files
dropZone.addEventListener('drop', handleDrop, false);

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = Array.from(dt.files);

    if (files.length === 0) return;

    if (isBatchMode || files.length > 1) {
        // Switch to batch mode if multiple files dropped
        if (!isBatchMode && files.length > 1) {
            setMode(true);
        }
        handleMultipleFiles(files);
    } else {
        handleFile(files[0]);
    }
}

// Handle file input change
fileInput.addEventListener('change', function() {
    if (this.files.length > 0) {
        handleFile(this.files[0]);
    }
});

fileInputMultiple.addEventListener('change', function() {
    if (this.files.length > 0) {
        handleMultipleFiles(Array.from(this.files));
    }
});

// Click drop zone to open file picker
// But not if clicking on the browse button label (which already triggers the input)
dropZone.addEventListener('click', function(e) {
    // Don't trigger if clicking on the label or browse button (they handle it themselves)
    if (e.target.closest('label') || e.target.classList.contains('browse-button')) {
        return;
    }
    if (isBatchMode) {
        fileInputMultiple.click();
    } else {
        fileInput.click();
    }
});

// Get file extension
function getFileExtension(filename) {
    return filename.split('.').pop().toLowerCase();
}

// Check if file is markdown
function isMarkdownFile(filename) {
    return markdownExtensions.includes(getFileExtension(filename));
}

// Check if file is docx
function isDocxFile(filename) {
    return docxExtensions.includes(getFileExtension(filename));
}

// Check if file is PDF
function isPdfFile(filename) {
    return pdfExtensions.includes(getFileExtension(filename));
}

// Check if file is zip
function isZipFile(filename) {
    return zipExtensions.includes(getFileExtension(filename));
}

// Check if file is convertible
function isConvertibleFile(filename) {
    return isMarkdownFile(filename) || isDocxFile(filename) || isPdfFile(filename);
}

// Get target formats for a given source file type
function getTargetFormats(filename) {
    if (isMarkdownFile(filename)) {
        return [
            { format: 'docx', label: 'Convert to DOCX', icon: '📄' },
            { format: 'pdf', label: 'Convert to PDF', icon: '📕' }
        ];
    } else if (isDocxFile(filename)) {
        return [
            { format: 'md', label: 'Convert to MD', icon: '📝' },
            { format: 'pdf', label: 'Convert to PDF', icon: '📕' }
        ];
    } else if (isPdfFile(filename)) {
        return [
            { format: 'md', label: 'Convert to MD', icon: '📝' },
            { format: 'docx', label: 'Convert to DOCX', icon: '📄' }
        ];
    }
    return [];
}

// Handle single file selection
function handleFile(file) {
    const fileExtension = getFileExtension(file.name);
    const allValidExtensions = [...markdownExtensions, ...docxExtensions, ...pdfExtensions];

    if (!allValidExtensions.includes(fileExtension)) {
        showError('Invalid file type. Please select a .md, .markdown, .txt, .docx, or .pdf file.');
        return;
    }

    // Validate file size (16MB for single file)
    const maxSize = 16 * 1024 * 1024;
    if (file.size > maxSize) {
        showError('File too large. Maximum size is 16MB.');
        return;
    }

    selectedFile = file;

    // Display file info
    fileNameDisplay.textContent = file.name;
    fileSizeDisplay.textContent = formatFileSize(file.size);

    // Create dual convert buttons based on source file type
    const targetFormats = getTargetFormats(file.name);
    convertButtonsContainer.innerHTML = '';

    targetFormats.forEach(target => {
        const btn = document.createElement('button');
        btn.className = `convert-button convert-to-${target.format}`;
        btn.innerHTML = `${target.icon} ${target.label}`;
        btn.addEventListener('click', () => convertFileWithFormat(target.format));
        convertButtonsContainer.appendChild(btn);
    });

    // Show file info and hide drop zone
    dropZone.classList.add('hidden');
    fileInfoContainer.classList.remove('hidden');
}

// Handle multiple files selection (batch mode)
function handleMultipleFiles(files) {
    // Validate total size (100MB for batch)
    const maxSize = 100 * 1024 * 1024;
    const totalSize = files.reduce((sum, f) => sum + f.size, 0);

    if (totalSize > maxSize) {
        showError('Total file size too large. Maximum is 100MB.');
        return;
    }

    // Filter and categorize files
    const convertibleFiles = [];
    const zipFiles = [];
    const skippedFiles = [];

    let mdFileCount = 0;
    let docxFileCount = 0;
    let pdfFileCount = 0;

    files.forEach(file => {
        if (isZipFile(file.name)) {
            zipFiles.push(file);
        } else if (isConvertibleFile(file.name)) {
            convertibleFiles.push(file);
            if (isMarkdownFile(file.name)) {
                mdFileCount++;
            } else if (isDocxFile(file.name)) {
                docxFileCount++;
            } else if (isPdfFile(file.name)) {
                pdfFileCount++;
            }
        } else {
            skippedFiles.push(file.name);
        }
    });

    // If only a zip file, handle it specially
    if (zipFiles.length === 1 && convertibleFiles.length === 0) {
        selectedFiles = zipFiles;
        batchTitle.textContent = 'ZIP Archive';
        batchCount.textContent = `${zipFiles[0].name} - Contents will be converted`;
        mdCount.textContent = '?';
        docxCount.textContent = '?';
        if (pdfCount) pdfCount.textContent = '?';

        // Clear and hide file list for zip
        batchFileList.innerHTML = '<p class="zip-note">ZIP contents will be extracted and converted</p>';

    } else if (convertibleFiles.length === 0 && zipFiles.length === 0) {
        showError('No convertible files found. Please select .md, .markdown, .txt, .docx, .pdf, or .zip files.');
        return;
    } else {
        selectedFiles = [...convertibleFiles, ...zipFiles];

        // Update batch info
        batchTitle.textContent = zipFiles.length > 0 ? 'Files and ZIP Archive' : 'Multiple Files';
        batchCount.textContent = `${selectedFiles.length} file${selectedFiles.length > 1 ? 's' : ''} ready for conversion`;
        mdCount.textContent = mdFileCount.toString();
        docxCount.textContent = docxFileCount.toString();
        if (pdfCount) pdfCount.textContent = pdfFileCount.toString();

        // Populate file list (show first 10)
        batchFileList.innerHTML = '';
        const displayFiles = selectedFiles.slice(0, 10);
        displayFiles.forEach(file => {
            const item = document.createElement('div');
            item.className = 'batch-file-item';

            let icon, direction;
            if (isZipFile(file.name)) {
                icon = '📦';
                direction = 'ZIP';
            } else if (isMarkdownFile(file.name)) {
                icon = '📝';
                direction = '→ .docx';
            } else if (isDocxFile(file.name)) {
                icon = '📄';
                direction = '→ .md';
            } else if (isPdfFile(file.name)) {
                icon = '📕';
                direction = '→ .md';
            } else {
                icon = '📋';
                direction = '';
            }

            item.innerHTML = `
                <span class="batch-file-icon">${icon}</span>
                <span class="batch-file-name">${file.name}</span>
                <span class="batch-file-direction">${direction}</span>
            `;
            batchFileList.appendChild(item);
        });

        if (selectedFiles.length > 10) {
            const more = document.createElement('p');
            more.className = 'batch-more';
            more.textContent = `... and ${selectedFiles.length - 10} more files`;
            batchFileList.appendChild(more);
        }
    }

    // Show batch info and hide drop zone
    dropZone.classList.add('hidden');
    batchFileInfo.classList.remove('hidden');
}

// Format file size
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';

    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));

    return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i];
}

// Cancel any ongoing conversion
function cancelConversion() {
    if (currentAbortController) {
        currentAbortController.abort();
        currentAbortController = null;
    }
}

// Fetch with timeout and retry support
async function fetchWithTimeout(url, options, timeout = CONFIG.REQUEST_TIMEOUT) {
    const controller = new AbortController();
    currentAbortController = controller;

    const timeoutId = setTimeout(() => controller.abort(), timeout);

    try {
        const response = await fetch(url, {
            ...options,
            signal: controller.signal
        });
        clearTimeout(timeoutId);
        return response;
    } catch (error) {
        clearTimeout(timeoutId);
        if (error.name === 'AbortError') {
            throw new Error('Request timed out or was cancelled. Please try again.');
        }
        throw error;
    } finally {
        currentAbortController = null;
    }
}

// Retry wrapper for fetch operations
async function fetchWithRetry(url, options, retries = CONFIG.RETRY_ATTEMPTS) {
    let lastError;

    for (let i = 0; i <= retries; i++) {
        try {
            return await fetchWithTimeout(url, options);
        } catch (error) {
            lastError = error;

            // Don't retry on user cancellation or certain errors
            if (error.message.includes('cancelled') ||
                error.message.includes('Invalid file') ||
                error.message.includes('too large')) {
                throw error;
            }

            // Wait before retry (exponential backoff)
            if (i < retries) {
                await new Promise(resolve => setTimeout(resolve, CONFIG.RETRY_DELAY * (i + 1)));
                updateProgressText(`Retrying... (attempt ${i + 2}/${retries + 1})`);
            }
        }
    }

    throw lastError;
}

// Update progress text with aria-live announcement
function updateProgressText(text) {
    progressText.textContent = text;
    // Announce to screen readers
    progressText.setAttribute('aria-live', 'polite');
}

// Convert file with specific target format
async function convertFileWithFormat(targetFormat) {
    if (!selectedFile) return;

    // Client-side file size validation
    if (selectedFile.size > CONFIG.MAX_SINGLE_FILE_SIZE) {
        showError(`File too large. Maximum size is ${formatFileSize(CONFIG.MAX_SINGLE_FILE_SIZE)}.`);
        return;
    }

    selectedTargetFormat = targetFormat;

    // Hide file info and show progress
    fileInfoContainer.classList.add('hidden');
    updateProgressText(`Converting to ${targetFormat.toUpperCase()}...`);
    progressContainer.classList.remove('hidden');

    const formData = new FormData();
    formData.append('file', selectedFile);

    try {
        // Add format query parameter
        const response = await fetchWithRetry(`/convert?format=${targetFormat}`, {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            let errorMessage = 'Conversion failed';
            try {
                const errorData = await response.json();
                errorMessage = errorData.error || errorMessage;
            } catch (jsonError) {
                // Response wasn't JSON, use status text
                errorMessage = response.statusText || errorMessage;
            }
            throw new Error(errorMessage);
        }

        // Download the file
        await downloadResponse(response, selectedFile.name, targetFormat);

        // Show success
        progressContainer.classList.add('hidden');
        successMessageText.textContent = 'Conversion successful! Your download should start automatically.';
        successMessage.classList.remove('hidden');

    } catch (error) {
        progressContainer.classList.add('hidden');
        showError(error.message);
    }
}

// Handle batch convert button click
batchConvertButton.addEventListener('click', convertBatch);

async function convertBatch() {
    if (selectedFiles.length === 0) return;

    // Client-side total size validation
    const totalSize = selectedFiles.reduce((sum, f) => sum + f.size, 0);
    if (totalSize > CONFIG.MAX_BATCH_SIZE) {
        showError(`Total file size too large. Maximum is ${formatFileSize(CONFIG.MAX_BATCH_SIZE)}.`);
        return;
    }

    // Validate individual file sizes
    for (const file of selectedFiles) {
        if (file.size > CONFIG.MAX_SINGLE_FILE_SIZE && !isZipFile(file.name)) {
            showError(`File "${file.name}" is too large. Maximum is ${formatFileSize(CONFIG.MAX_SINGLE_FILE_SIZE)} per file.`);
            return;
        }
    }

    // Hide batch info and show progress
    batchFileInfo.classList.add('hidden');
    updateProgressText(`Converting ${selectedFiles.length} file${selectedFiles.length > 1 ? 's' : ''}...`);
    progressContainer.classList.remove('hidden');

    const formData = new FormData();

    // Check if it's a single zip file
    if (selectedFiles.length === 1 && isZipFile(selectedFiles[0].name)) {
        formData.append('file', selectedFiles[0]);
    } else {
        // Multiple files
        selectedFiles.forEach(file => {
            formData.append('files', file);
        });
    }

    try {
        const response = await fetchWithRetry('/convert-batch', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            let errorMessage = 'Batch conversion failed';
            try {
                const errorData = await response.json();
                errorMessage = errorData.error || errorMessage;
            } catch (jsonError) {
                // Response wasn't JSON, use status text
                errorMessage = response.statusText || errorMessage;
            }
            throw new Error(errorMessage);
        }

        // Download the response
        const contentType = response.headers.get('Content-Type');
        const isZip = contentType && contentType.includes('application/zip');

        await downloadResponse(response, isZip ? 'converted_files.zip' : null);

        // Show success
        progressContainer.classList.add('hidden');
        const fileCount = selectedFiles.length;
        successMessageText.textContent = isZip
            ? `Converted ${fileCount} files! ZIP download should start automatically.`
            : 'Conversion successful! Your download should start automatically.';
        successMessage.classList.remove('hidden');

    } catch (error) {
        progressContainer.classList.add('hidden');
        showError(error.message);
    }
}

// Download response helper
async function downloadResponse(response, fallbackName, targetFormat) {
    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.style.display = 'none';
    a.href = url;

    // Get filename from Content-Disposition header or use fallback
    const contentDisposition = response.headers.get('Content-Disposition');
    let filename = fallbackName || 'converted';

    if (contentDisposition) {
        const filenameMatch = contentDisposition.match(/filename="?([^"]+)"?/);
        if (filenameMatch) {
            filename = filenameMatch[1];
        }
    } else if (fallbackName && selectedFile && targetFormat) {
        const originalName = selectedFile.name.replace(/\.[^/.]+$/, '');
        filename = originalName + '.' + targetFormat;
    }

    a.download = filename;
    document.body.appendChild(a);
    a.click();

    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
}

// Show error message
function showError(message) {
    errorMessageText.textContent = message;
    errorMessage.classList.remove('hidden');
}

// Reset the form
function resetForm() {
    // Cancel any ongoing conversion
    cancelConversion();

    selectedFile = null;
    selectedFiles = [];
    selectedTargetFormat = null;

    // Reset file inputs properly (create new elements to ensure change event fires)
    const newFileInput = fileInput.cloneNode(true);
    const newFileInputMultiple = fileInputMultiple.cloneNode(true);

    fileInput.parentNode.replaceChild(newFileInput, fileInput);
    fileInputMultiple.parentNode.replaceChild(newFileInputMultiple, fileInputMultiple);

    // Update references and re-attach event listeners
    const updatedFileInput = document.getElementById('file-input');
    const updatedFileInputMultiple = document.getElementById('file-input-multiple');

    updatedFileInput.addEventListener('change', function() {
        if (this.files.length > 0) {
            handleFile(this.files[0]);
        }
    });

    updatedFileInputMultiple.addEventListener('change', function() {
        if (this.files.length > 0) {
            handleMultipleFiles(Array.from(this.files));
        }
    });

    // Clear convert buttons
    if (convertButtonsContainer) {
        convertButtonsContainer.innerHTML = '';
    }

    // Hide all messages and containers
    fileInfoContainer.classList.add('hidden');
    batchFileInfo.classList.add('hidden');
    progressContainer.classList.add('hidden');
    successMessage.classList.add('hidden');
    errorMessage.classList.add('hidden');

    // Show drop zone
    dropZone.classList.remove('hidden');
}

// Reset button handlers
resetButton.addEventListener('click', resetForm);
errorResetButton.addEventListener('click', resetForm);

// Cancel button handler
const cancelButton = document.getElementById('cancel-button');
if (cancelButton) {
    cancelButton.addEventListener('click', function() {
        cancelConversion();
        progressContainer.classList.add('hidden');
        showError('Conversion cancelled.');
    });
}

// Keyboard navigation for drop zone
dropZone.addEventListener('keydown', function(e) {
    if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        if (isBatchMode) {
            document.getElementById('file-input-multiple').click();
        } else {
            document.getElementById('file-input').click();
        }
    }
});

// Update aria-pressed on mode toggle
function updateModeAriaStates() {
    singleModeBtn.setAttribute('aria-pressed', !isBatchMode);
    batchModeBtn.setAttribute('aria-pressed', isBatchMode);
}
