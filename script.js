// Global variables
let currentConverter = null;
let currentFile = null;
const { jsPDF } = window.jspdf;

// Image Editor variables
let editorCanvas = document.getElementById('editor-canvas');
let editorCtx = editorCanvas.getContext('2d');
let originalImageData = null;
let currentFilter = null;

// Tab functionality
document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
        document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        
        tab.classList.add('active');
        document.getElementById(`${tab.dataset.tab}-tab`).classList.add('active');
    });
});

// Converter card functionality
document.querySelectorAll('.converter-card').forEach(card => {
    card.addEventListener('click', () => {
        currentConverter = card.dataset.converter;
        document.getElementById('converter-ui').classList.add('active');
        document.getElementById('file-input').value = '';
        document.getElementById('result-container').style.display = 'none';
        document.getElementById('progress-container').style.display = 'none';
        document.getElementById('convert-btn').style.display = 'none';
        
        // Set converter title
        document.getElementById('converter-title').textContent = card.querySelector('h3').textContent;
        
        // Set supported formats
        let formats = '';
        switch(currentConverter) {
            case 'word-to-pdf':
            case 'pdf-to-word':
            case 'word-to-excel':
            case 'word-to-image':
            case 'image-to-word':
                formats = 'DOC, DOCX';
                break;
            case 'pdf-edit':
            case 'pdf-merge':
            case 'pdf-compress':
            case 'pdf-split':
            case 'pdf-repair':
            case 'image-to-pdf':
                formats = 'PDF';
                break;
            case 'image-compression':
            case 'image-merge':
            case 'image-editor':
            case 'image-split':
                formats = 'JPG, PNG, GIF, BMP';
                break;
            case 'word-to-ppt':
            case 'ppt-to-word':
            case 'ppt-to-pdf':
            case 'pdf-to-ppt':
                formats = 'PPT, PPTX';
                break;
        }
        
        if (currentConverter === 'pdf-merge' || currentConverter === 'image-merge') {
            document.getElementById('file-instructions').textContent = 'Drag & drop multiple files here or click to browse';
            document.getElementById('file-input').setAttribute('multiple', '');
        } else {
            document.getElementById('file-instructions').textContent = 'Drag & drop your file here or click to browse';
            document.getElementById('file-input').removeAttribute('multiple');
        }
        
        document.getElementById('supported-formats').textContent = `Supported formats: ${formats}`;
        
        // Set options
        const optionsContainer = document.getElementById('conversion-options');
        optionsContainer.innerHTML = '';
        optionsContainer.classList.remove('active');
        
        // Show content for selected tool
        document.querySelectorAll('.tool-content').forEach(tc => tc.style.display = 'none');
        const content = document.getElementById(`${currentConverter}-info`);
        if (content) content.style.display = 'block';
        
        // Scroll to converter UI
        document.getElementById('converter-ui').scrollIntoView({ behavior: 'smooth' });
    });
});

// File input functionality
document.getElementById('file-drop-area').addEventListener('click', () => {
    document.getElementById('file-input').click();
});

document.getElementById('file-input').addEventListener('change', handleFileSelect);

// Drag and drop functionality
const fileDropArea = document.getElementById('file-drop-area');

['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    fileDropArea.addEventListener(eventName, preventDefaults, false);
});

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

['dragenter', 'dragover'].forEach(eventName => {
    fileDropArea.addEventListener(eventName, highlight, false);
});

['dragleave', 'drop'].forEach(eventName => {
    fileDropArea.addEventListener(eventName, unhighlight, false);
});

function highlight() {
    fileDropArea.style.borderColor = 'var(--primary)';
    fileDropArea.style.backgroundColor = 'rgba(162, 155, 254, 0.1)';
}

function unhighlight() {
    fileDropArea.style.borderColor = 'var(--secondary)';
    fileDropArea.style.backgroundColor = '';
}

fileDropArea.addEventListener('drop', handleDrop, false);

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    
    if (files.length > 0) {
        handleFiles(files);
    }
}

function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        handleFiles(files);
    }
}

function handleFiles(files) {
    if (currentConverter === null) return;
    
    // For merge operations, we handle multiple files
    if ((currentConverter === 'pdf-merge' || currentConverter === 'image-merge') && files.length > 1) {
        currentFile = files;
        prepareForConversion();
        return;
    }
    
    // For other operations, we only handle the first file
    currentFile = files[0];
    prepareForConversion();
}

function prepareForConversion() {
    document.getElementById('convert-btn').style.display = 'block';
    
    // Show file info
    if (currentFile instanceof FileList || Array.isArray(currentFile)) {
        document.getElementById('file-instructions').textContent = `${currentFile.length} files selected`;
    } else {
        document.getElementById('file-instructions').textContent = currentFile.name;
    }
    
    // Show options if available
    const optionsContainer = document.getElementById('conversion-options');
    optionsContainer.innerHTML = '';
    
    switch(currentConverter) {
        case 'pdf-compress':
            optionsContainer.innerHTML = `
                <div class="option-group">
                    <label for="compression-level">Compression Level</label>
                    <select id="compression-level">
                        <option value="low">Low (Smaller size reduction, better quality)</option>
                        <option value="medium" selected>Medium (Balance between size and quality)</option>
                        <option value="high">High (Maximum size reduction, lower quality)</option>
                    </select>
                </div>
            `;
            break;
            
        case 'pdf-split':
            optionsContainer.innerHTML = `
                <div class="option-group">
                    <label for="split-method">Split Method</label>
                    <select id="split-method">
                        <option value="range">Split by page range</option>
                        <option value="every">Split every X pages</option>
                    </select>
                </div>
                <div class="option-group" id="range-options">
                    <label for="page-range">Page Range (e.g., 1-3,5,7-10)</label>
                    <input type="text" id="page-range" placeholder="1-3,5,7-10">
                </div>
                <div class="option-group" id="every-options" style="display: none;">
                    <label for="every-pages">Split every X pages</label>
                    <input type="number" id="every-pages" min="1" value="1">
                </div>
            `;
            
            document.getElementById('split-method').addEventListener('change', function() {
                document.getElementById('range-options').style.display = 
                    this.value === 'range' ? 'block' : 'none';
                document.getElementById('every-options').style.display = 
                    this.value === 'every' ? 'block' : 'none';
            });
            break;
            
        case 'image-compression':
            optionsContainer.innerHTML = `
                <div class="option-group">
                    <label for="image-quality">Image Quality (0-100)</label>
                    <input type="range" id="image-quality" min="0" max="100" value="80">
                    <span id="quality-value">80%</span>
                </div>
                <div class="option-group">
                    <label for="image-width">Resize Width (px, 0 to keep original)</label>
                    <input type="number" id="image-width" min="0" value="0">
                </div>
                <div class="option-group">
                    <label for="image-height">Resize Height (px, 0 to keep original)</label>
                    <input type="number" id="image-height" min="0" value="0">
                </div>
            `;
            
            document.getElementById('image-quality').addEventListener('input', function() {
                document.getElementById('quality-value').textContent = `${this.value}%`;
            });
            break;
            
        case 'image-merge':
            optionsContainer.innerHTML = `
                <div class="option-group">
                    <label for="merge-direction">Merge Direction</label>
                    <select id="merge-direction">
                        <option value="vertical">Vertical (Top to Bottom)</option>
                        <option value="horizontal">Horizontal (Left to Right)</option>
                    </select>
                </div>
                <div class="option-group">
                    <label for="merge-spacing">Spacing Between Images (px)</label>
                    <input type="number" id="merge-spacing" min="0" value="0">
                </div>
                <div class="option-group">
                    <label for="merge-background">Background Color</label>
                    <input type="color" id="merge-background" value="#ffffff">
                </div>
            `;
            break;
    }
    
    if (optionsContainer.innerHTML !== '') {
        optionsContainer.classList.add('active');
    }
}

// Convert button functionality
document.getElementById('convert-btn').addEventListener('click', startConversion);

function startConversion() {
    if (!currentFile || !currentConverter) return;
    
    document.getElementById('convert-btn').style.display = 'none';
    document.getElementById('progress-container').style.display = 'block';
    document.getElementById('progress-bar').style.width = '0%';
    document.getElementById('progress-text').textContent = '0%';
    
    // Simulate progress for demonstration
    let progress = 0;
    const progressInterval = setInterval(() => {
        progress += 5;
        if (progress > 100) progress = 100;
        document.getElementById('progress-bar').style.width = `${progress}%`;
        document.getElementById('progress-text').textContent = `${progress}%`;
        
        if (progress === 100) {
            clearInterval(progressInterval);
            performConversion();
        }
    }, 200);
}

async function performConversion() {
    try {
        let result;
        
        switch(currentConverter) {
            case 'word-to-pdf':
                result = await convertWordToPdf(currentFile);
                break;
            case 'word-to-excel':
                result = await convertWordToExcel(currentFile);
                break;
            case 'word-to-image':
                result = await convertWordToImage(currentFile);
                break;
            case 'pdf-to-word':
                result = await convertPdfToWord(currentFile);
                break;
            case 'image-to-word':
                result = await convertImageToWord(currentFile);
                break;
            case 'pdf-edit':
                result = await editPdf(currentFile);
                break;
            case 'pdf-merge':
                result = await mergePdfs(currentFile);
                break;
            case 'pdf-compress':
                result = await compressPdf(currentFile);
                break;
            case 'pdf-split':
                result = await splitPdf(currentFile);
                break;
            case 'pdf-repair':
                result = await repairPdf(currentFile);
                break;
            case 'image-to-pdf':
                result = await convertImageToPdf(currentFile);
                break;
            case 'image-compression':
                result = await compressImage(currentFile);
                break;
            case 'image-merge':
                result = await mergeImages(currentFile);
                break;
            case 'image-editor':
                result = await editImage(currentFile);
                break;
            case 'image-split':
                result = await splitImage(currentFile);
                break;
            case 'word-to-ppt':
                result = await convertWordToPpt(currentFile);
                break;
            case 'ppt-to-word':
                result = await convertPptToWord(currentFile);
                break;
            case 'ppt-to-pdf':
                result = await convertPptToPdf(currentFile);
                break;
            case 'pdf-to-ppt':
                result = await convertPdfToPpt(currentFile);
                break;
            default:
                throw new Error('Unsupported conversion type');
        }
        
        showResult(result);
    } catch (error) {
        console.error('Conversion error:', error);
        alert(`Conversion failed: ${error.message}`);
        resetConverterUI();
    }
}

function showResult(result) {
    document.getElementById('progress-container').style.display = 'none';
    document.getElementById('result-container').style.display = 'block';
    
    const downloadBtn = document.getElementById('download-btn');
    downloadBtn.href = URL.createObjectURL(result.file);
    downloadBtn.download = result.filename;
    
    // Scroll to result
    document.getElementById('result-container').scrollIntoView({ behavior: 'smooth' });
}

// New conversion button
document.getElementById('new-conversion').addEventListener('click', resetConverterUI);

function resetConverterUI() {
    currentFile = null;
    document.getElementById('file-input').value = '';
    document.getElementById('file-instructions').textContent = 'Drag & drop your file here or click to browse';
    document.getElementById('result-container').style.display = 'none';
    document.getElementById('convert-btn').style.display = 'none';
    document.getElementById('conversion-options').classList.remove('active');
}

// Image Editor Functions
async function editImage(file) {
    return new Promise((resolve) => {
        // Create image from file
        const img = new Image();
        const url = URL.createObjectURL(file);
        
        img.onload = function() {
            // Set canvas dimensions
            editorCanvas.width = img.width;
            editorCanvas.height = img.height;
            
            // Draw image on canvas
            editorCtx.drawImage(img, 0, 0);
            
            // Store original image data
            originalImageData = editorCtx.getImageData(0, 0, editorCanvas.width, editorCanvas.height);
            
            // Show editor modal
            document.getElementById('image-editor-modal').classList.add('active');
            
            // Set up editor buttons
            setupImageEditor(resolve, file.name);
        };
        
        img.src = url;
    });
}

function setupImageEditor(resolveCallback, originalFilename) {
    // Reset range slider
    document.getElementById('editor-range').value = 0;
    
    // Set up tool buttons
    document.getElementById('brightness-btn').onclick = () => {
        currentFilter = 'brightness';
        updateImageFilter();
    };
    
    document.getElementById('contrast-btn').onclick = () => {
        currentFilter = 'contrast';
        updateImageFilter();
    };
    
    document.getElementById('saturation-btn').onclick = () => {
        currentFilter = 'saturation';
        updateImageFilter();
    };
    
    document.getElementById('grayscale-btn').onclick = () => {
        currentFilter = 'grayscale';
        updateImageFilter();
    };
    
    document.getElementById('sepia-btn').onclick = () => {
        currentFilter = 'sepia';
        updateImageFilter();
    };
    
    document.getElementById('reset-btn').onclick = resetImage;
    
    // Set up range slider
    document.getElementById('editor-range').oninput = updateImageFilter;
    
    // Set up action buttons
    document.getElementById('cancel-edit').onclick = () => {
        document.getElementById('image-editor-modal').classList.remove('active');
        resolveCallback(null); // Return null to indicate cancellation
    };
    
    document.getElementById('apply-edit').onclick = () => {
        // Convert canvas to blob
        editorCanvas.toBlob((blob) => {
            document.getElementById('image-editor-modal').classList.remove('active');
            resolveCallback({
                file: blob,
                filename: 'edited_' + originalFilename
            });
        }, 'image/jpeg', 0.9);
    };
}

function resetImage() {
    if (originalImageData) {
        editorCtx.putImageData(originalImageData, 0, 0);
        currentFilter = null;
        document.getElementById('editor-range').value = 0;
    }
}

function updateImageFilter() {
    if (!originalImageData) return;
    
    // Reset to original image
    editorCtx.putImageData(originalImageData, 0, 0);
    
    const value = parseInt(document.getElementById('editor-range').value);
    
    switch(currentFilter) {
        case 'brightness':
            applyBrightness(value);
            break;
        case 'contrast':
            applyContrast(value);
            break;
        case 'saturation':
            applySaturation(value);
            break;
        case 'grayscale':
            applyGrayscale();
            break;
        case 'sepia':
            applySepia();
            break;
    }
}

function applyBrightness(value) {
    const imageData = editorCtx.getImageData(0, 0, editorCanvas.width, editorCanvas.height);
    const data = imageData.data;
    const factor = (value + 100) / 100;
    
    for (let i = 0; i < data.length; i += 4) {
        data[i] = data[i] * factor;     // R
        data[i + 1] = data[i + 1] * factor; // G
        data[i + 2] = data[i + 2] * factor; // B
    }
    
    editorCtx.putImageData(imageData, 0, 0);
}

function applyContrast(value) {
    const imageData = editorCtx.getImageData(0, 0, editorCanvas.width, editorCanvas.height);
    const data = imageData.data;
    const factor = (value + 100) / 100;
    const intercept = 128 * (1 - factor);
    
    for (let i = 0; i < data.length; i += 4) {
        data[i] = data[i] * factor + intercept;     // R
        data[i + 1] = data[i + 1] * factor + intercept; // G
        data[i + 2] = data[i + 2] * factor + intercept; // B
    }
    
    editorCtx.putImageData(imageData, 0, 0);
}

function applySaturation(value) {
    const imageData = editorCtx.getImageData(0, 0, editorCanvas.width, editorCanvas.height);
    const data = imageData.data;
    const factor = (value + 100) / 100;
    
    for (let i = 0; i < data.length; i += 4) {
        const r = data[i];
        const g = data[i + 1];
        const b = data[i + 2];
        const gray = 0.2989 * r + 0.5870 * g + 0.1140 * b; // Grayscale value
        
        data[i] = gray + (r - gray) * factor;     // R
        data[i + 1] = gray + (g - gray) * factor; // G
        data[i + 2] = gray + (b - gray) * factor; // B
    }
    
    editorCtx.putImageData(imageData, 0, 0);
}

function applyGrayscale() {
    const imageData = editorCtx.getImageData(0, 0, editorCanvas.width, editorCanvas.height);
    const data = imageData.data;
    
    for (let i = 0; i < data.length; i += 4) {
        const avg = 0.2989 * data[i] + 0.5870 * data[i + 1] + 0.1140 * data[i + 2];
        data[i] = data[i + 1] = data[i + 2] = avg;
    }
    
    editorCtx.putImageData(imageData, 0, 0);
}

function applySepia() {
    const imageData = editorCtx.getImageData(0, 0, editorCanvas.width, editorCanvas.height);
    const data = imageData.data;
    
    for (let i = 0; i < data.length; i += 4) {
        const r = data[i];
        const g = data[i + 1];
        const b = data[i + 2];
        
        data[i] = Math.min(255, (r * 0.393) + (g * 0.769) + (b * 0.189));
        data[i + 1] = Math.min(255, (r * 0.349) + (g * 0.686) + (b * 0.168));
        data[i + 2] = Math.min(255, (r * 0.272) + (g * 0.534) + (b * 0.131));
    }
    
    editorCtx.putImageData(imageData, 0, 0);
}

// Conversion Functions
async function convertWordToPdf(file) {
    const formData = new FormData();
    formData.append('file', file);

    const response = await fetch('https://api.fileconvert.co.in/convert/word-to-pdf', {
        method: 'POST',
        body: formData
    });

    if (!response.ok) {
        throw new Error('Conversion failed.');
    }

    const blob = await response.blob();

    return {
        file: blob,
        filename: file.name.replace(/\.[^/.]+$/, '') + '.pdf'
    };
}



async function convertWordToExcel(file) {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = result.value;
    const tables = tempDiv.querySelectorAll('table');
    
    if (tables.length === 0) {
        throw new Error('No tables found in the Word document');
    }
    
    const wb = XLSX.utils.book_new();
    
    tables.forEach((table, index) => {
        const ws = XLSX.utils.table_to_sheet(table);
        XLSX.utils.book_append_sheet(wb, ws, `Table ${index + 1}`);
    });
    
    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const excelBlob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    return {
        file: excelBlob,
        filename: file.name.replace(/\.[^/.]+$/, '') + '.xlsx'
    };
}

async function convertWordToImage(file) {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = result.value;
    document.body.appendChild(tempDiv);
    
    const canvas = await html2canvas(tempDiv);
    document.body.removeChild(tempDiv);
    
    return new Promise((resolve) => {
        canvas.toBlob((blob) => {
            resolve({
                file: blob,
                filename: file.name.replace(/\.[^/.]+$/, '') + '.png'
            });
        }, 'image/png');
    });
}

async function convertPdfToWord(file) {
    const formData = new FormData();
    formData.append('file', file);

    const response = await fetch('https://api.fileconvert.co.in/convert/pdf-to-word', {
        method: 'POST',
        body: formData
    });

    if (!response.ok) throw new Error('Conversion failed.');

    const blob = await response.blob();

    return {
        file: blob,
        filename: file.name.replace(/\.[^/.]+$/, '') + '.docx'
    };
}

async function convertImageToWord(file) {
    const { Document, Paragraph, Packer, ImageRun } = docx;
    const arrayBuffer = await file.arrayBuffer();
    
    const doc = new Document({
        sections: [{
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new ImageRun({
                            data: arrayBuffer,
                            transformation: {
                                width: 400,
                                height: 300,
                            },
                        })
                    ]
                })
            ]
        }]
    });
    
    const docBuffer = await Packer.toBuffer(doc);
    const docBlob = new Blob([docBuffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    
    return {
        file: docBlob,
        filename: file.name.replace(/\.[^/.]+$/, '') + '.docx'
    };
}

async function editPdf(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdfDoc = await PDFLib.PDFDocument.load(arrayBuffer);
    
    const pages = pdfDoc.getPages();
    if (pages.length === 0) {
        throw new Error('PDF has no pages');
    }
    
    const firstPage = pages[0];
    
    firstPage.drawText('Edited with Easy File Conversion', {
        x: 50,
        y: firstPage.getHeight() - 50,
        size: 30,
        color: PDFLib.rgb(0.95, 0.1, 0.1),
    });
    
    const pdfBytes = await pdfDoc.save();
    const pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
    
    return {
        file: pdfBlob,
        filename: 'edited_' + file.name
    };
}

async function mergePdfs(files) {
    const pdfDoc = await PDFLib.PDFDocument.create();
    
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const arrayBuffer = await file.arrayBuffer();
        const donorPdfDoc = await PDFLib.PDFDocument.load(arrayBuffer);
        
        const pages = await pdfDoc.copyPages(donorPdfDoc, donorPdfDoc.getPageIndices());
        pages.forEach(page => pdfDoc.addPage(page));
    }
    
    const pdfBytes = await pdfDoc.save();
    const pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
    
    return {
        file: pdfBlob,
        filename: 'merged_documents.pdf'
    };
}

async function compressPdf(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdfDoc = await PDFLib.PDFDocument.load(arrayBuffer);
    
    const compressionLevel = document.getElementById('compression-level').value;
    
    const saveOptions = {
        useObjectStreams: true,
    };
    
    switch(compressionLevel) {
        case 'low':
            saveOptions.useCompression = true;
            break;
        case 'medium':
            saveOptions.useCompression = true;
            saveOptions.keepExistingFields = false;
            break;
        case 'high':
            saveOptions.useCompression = true;
            saveOptions.keepExistingFields = false;
            saveOptions.keepIndirectObjects = false;
            break;
    }
    
    const pdfBytes = await pdfDoc.save(saveOptions);
    const pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
    
    return {
        file: pdfBlob,
        filename: 'compressed_' + file.name
    };
}

async function splitPdf(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdfDoc = await PDFLib.PDFDocument.load(arrayBuffer);
    
    const splitMethod = document.getElementById('split-method').value;
    let pageRanges = [];
    
    if (splitMethod === 'range') {
        const rangeText = document.getElementById('page-range').value;
        pageRanges = parsePageRanges(rangeText, pdfDoc.getPageCount());
    } else {
        const everyX = parseInt(document.getElementById('every-pages').value);
        const pageCount = pdfDoc.getPageCount();
        
        for (let i = 0; i < pageCount; i += everyX) {
            const end = Math.min(i + everyX, pageCount);
            pageRanges.push({ start: i + 1, end: end });
        }
    }
    
    if (pageRanges.length === 0) {
        throw new Error('No valid page ranges specified');
    }
    
    const JSZip = await loadJSZip();
    const zip = new JSZip();
    
    for (let i = 0; i < pageRanges.length; i++) {
        const range = pageRanges[i];
        const newPdfDoc = await PDFLib.PDFDocument.create();
        
        const pagesToCopy = [];
        for (let j = range.start; j <= range.end; j++) {
            pagesToCopy.push(j - 1);
        }
        
        const pages = await newPdfDoc.copyPages(pdfDoc, pagesToCopy);
        pages.forEach(page => newPdfDoc.addPage(page));
        
        const pdfBytes = await newPdfDoc.save();
        const fileName = file.name.replace(/\.[^/.]+$/, '') + `_part${i + 1}.pdf`;
        zip.file(fileName, pdfBytes);
    }
    
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    
    return {
        file: zipBlob,
        filename: 'split_' + file.name.replace(/\.[^/.]+$/, '') + '.zip'
    };
}

function parsePageRanges(rangeText, maxPages) {
    const ranges = [];
    const parts = rangeText.split(',');
    
    for (const part of parts) {
        if (part.includes('-')) {
            const [start, end] = part.split('-').map(num => parseInt(num.trim()));
            if (!isNaN(start) && !isNaN(end)) {
                ranges.push({
                    start: Math.max(1, Math.min(start, maxPages)),
                    end: Math.max(1, Math.min(end, maxPages))
                });
            }
        } else {
            const page = parseInt(part.trim());
            if (!isNaN(page)) {
                const validPage = Math.max(1, Math.min(page, maxPages));
                ranges.push({
                    start: validPage,
                    end: validPage
                });
            }
        }
    }
    
    return ranges;
}

async function repairPdf(file) {
    try {
        const arrayBuffer = await file.arrayBuffer();
        const pdfDoc = await PDFLib.PDFDocument.load(arrayBuffer, { ignoreEncryption: true });
        
        const pdfBytes = await pdfDoc.save();
        const pdfBlob = new Blob([pdfBytes], { type: 'application/pdf' });
        
        return {
            file: pdfBlob,
            filename: 'repaired_' + file.name
        };
    } catch (error) {
        throw new Error('Unable to repair the PDF file. It may be too severely damaged.');
    }
}

async function convertImageToPdf(file) {
    const img = await createImageBitmap(file);
    const canvas = document.createElement('canvas');
    canvas.width = img.width;
    canvas.height = img.height;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0);
    
    const pdf = new jsPDF({
        orientation: img.width > img.height ? 'landscape' : 'portrait',
        unit: 'mm'
    });
    
    const imgData = canvas.toDataURL('image/jpeg');
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = pdf.internal.pageSize.getHeight();
    
    const imgRatio = img.width / img.height;
    const pdfRatio = pdfWidth / pdfHeight;
    
    let width, height;
    if (imgRatio > pdfRatio) {
        width = pdfWidth;
        height = pdfWidth / imgRatio;
    } else {
        height = pdfHeight;
        width = pdfHeight * imgRatio;
    }
    
    pdf.addImage(imgData, 'JPEG', 0, 0, width, height);
    const pdfBlob = pdf.output('blob');
    
    return {
        file: pdfBlob,
        filename: file.name.replace(/\.[^/.]+$/, '') + '.pdf'
    };
}

async function compressImage(file) {
    const quality = parseInt(document.getElementById('image-quality').value) / 100;
    const width = parseInt(document.getElementById('image-width').value) || undefined;
    const height = parseInt(document.getElementById('image-height').value) || undefined;
    
    return new Promise((resolve) => {
        new Compressor(file, {
            quality,
            width,
            height,
            success(result) {
                resolve({
                    file: result,
                    filename: 'compressed_' + file.name
                });
            },
            error(err) {
                throw new Error('Image compression failed: ' + err.message);
            }
        });
    });
}

async function mergeImages(files) {
    const direction = document.getElementById('merge-direction').value;
    const spacing = parseInt(document.getElementById('merge-spacing').value);
    const bgColor = document.getElementById('merge-background').value;
    
    const images = [];
    for (let i = 0; i < files.length; i++) {
        const img = await createImageBitmap(files[i]);
        images.push(img);
    }
    
    let canvasWidth = 0;
    let canvasHeight = 0;
    
    if (direction === 'horizontal') {
        canvasWidth = images.reduce((sum, img) => sum + img.width, 0) + (spacing * (images.length - 1));
        canvasHeight = Math.max(...images.map(img => img.height));
    } else {
        canvasWidth = Math.max(...images.map(img => img.width));
        canvasHeight = images.reduce((sum, img) => sum + img.height, 0) + (spacing * (images.length - 1));
    }
    
    const canvas = document.createElement('canvas');
    canvas.width = canvasWidth;
    canvas.height = canvasHeight;
    const ctx = canvas.getContext('2d');
    
    ctx.fillStyle = bgColor;
    ctx.fillRect(0, 0, canvasWidth, canvasHeight);
    
    let x = 0;
    let y = 0;
    
    for (const img of images) {
        ctx.drawImage(img, x, y);
        
        if (direction === 'horizontal') {
            x += img.width + spacing;
        } else {
            y += img.height + spacing;
        }
    }
    
    return new Promise((resolve) => {
        canvas.toBlob((blob) => {
            resolve({
                file: blob,
                filename: 'merged_images.png'
            });
        }, 'image/png');
    });
}

async function splitImage(file) {
    const img = await createImageBitmap(file);
    const canvas = document.createElement('canvas');
    canvas.width = img.width;
    canvas.height = img.height;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(img, 0, 0);
    
    const halfWidth = img.width / 2;
    const halfHeight = img.height / 2;
    
    const parts = [
        { x: 0, y: 0, width: halfWidth, height: halfHeight },
        { x: halfWidth, y: 0, width: halfWidth, height: halfHeight },
        { x: 0, y: halfHeight, width: halfWidth, height: halfHeight },
        { x: halfWidth, y: halfHeight, width: halfWidth, height: halfHeight }
    ];
    
    const JSZip = await loadJSZip();
    const zip = new JSZip();
    
    for (let i = 0; i < parts.length; i++) {
        const part = parts[i];
        const partCanvas = document.createElement('canvas');
        partCanvas.width = part.width;
        partCanvas.height = part.height;
        const partCtx = partCanvas.getContext('2d');
        
        partCtx.drawImage(
            canvas,
            part.x, part.y, part.width, part.height,
            0, 0, part.width, part.height
        );
        
        const blob = await new Promise(resolve => 
            partCanvas.toBlob(resolve, 'image/jpeg', 0.9)
        );
        
        const fileName = file.name.replace(/\.[^/.]+$/, '') + `_part${i + 1}.jpg`;
        zip.file(fileName, blob);
    }
    
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    
    return {
        file: zipBlob,
        filename: 'split_' + file.name.replace(/\.[^/.]+$/, '') + '.zip'
    };
}

async function convertWordToPpt(file) {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    
    const PptxGenJS = await loadPptxGenJS();
    const pptx = new PptxGenJS();
    
    const slide = pptx.addSlide();
    slide.addText(result.value, { x: 0.5, y: 0.5, w: 9, h: 6.5, fontSize: 18 });
    
    const pptBuffer = await pptx.write({ outputType: 'arraybuffer' });
    const pptBlob = new Blob([pptBuffer], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    
    return {
        file: pptBlob,
        filename: file.name.replace(/\.[^/.]+$/, '') + '.pptx'
    };
}

async function convertPptToWord(file) {
    const { Document, Paragraph, Packer, TextRun } = docx;
    
    const doc = new Document({
        sections: [{
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun('PPT content would be extracted here. In a real implementation, you would use a library to extract text from PowerPoint slides.')
                    ]
                })
            ]
        }]
    });
    
    const docBuffer = await Packer.toBuffer(doc);
    const docBlob = new Blob([docBuffer], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    
    return {
        file: docBlob,
        filename: file.name.replace(/\.[^/.]+$/, '') + '.docx'
    };
}

async function convertPptToPdf(file) {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF();
    
    pdf.text('PPT slides would be converted to PDF here. In a real implementation, you would use a library to convert PowerPoint slides to PDF.', 10, 10);
    
    const pdfBlob = pdf.output('blob');
    return {
        file: pdfBlob,
        filename: file.name.replace(/\.[^/.]+$/, '') + '.pdf'
    };
}

async function convertPdfToPpt(file) {
    const PptxGenJS = await loadPptxGenJS();
    const pptx = new PptxGenJS();
    
    const slide = pptx.addSlide();
    slide.addText('PDF pages would be converted to PPT slides here. In a real implementation, you would use a library to convert PDF pages to images and add them as slides.', { x: 0.5, y: 0.5, w: 9, h: 6.5, fontSize: 18 });
    
    const pptBuffer = await pptx.write({ outputType: 'arraybuffer' });
    const pptBlob = new Blob([pptBuffer], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
    
    return {
        file: pptBlob,
        filename: file.name.replace(/\.[^/.]+$/, '') + '.pptx'
    };
}

// Helper function to dynamically load JSZip
function loadJSZip() {
    return new Promise((resolve) => {
        if (window.JSZip) {
            resolve(window.JSZip);
        } else {
            const script = document.createElement('script');
            script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.7.1/jszip.min.js';
            script.onload = () => resolve(window.JSZip);
            document.head.appendChild(script);
        }
    });
}

// Helper function to dynamically load pptxgenjs
function loadPptxGenJS() {
    return new Promise((resolve) => {
        if (window.PptxGenJS) {
            resolve(window.PptxGenJS);
        } else {
            const script = document.createElement('script');
            script.src = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.7.0/dist/pptxgen.bundle.js';
            script.onload = () => resolve(window.PptxGenJS);
            document.head.appendChild(script);
        }
    });

}
