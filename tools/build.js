#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

console.log('🏗️  Building static version for shared hosting...');

// Create dist directory
const distDir = path.join(__dirname, '..', 'dist');
if (fs.existsSync(distDir)) {
    fs.rmSync(distDir, { recursive: true });
}
fs.mkdirSync(distDir, { recursive: true });

// Copy client assets
const clientDir = path.join(__dirname, '..', 'src', 'client');
const assetsDir = path.join(distDir, 'assets');
fs.mkdirSync(assetsDir, { recursive: true });

// Copy CSS files with path corrections
const stylesDir = path.join(clientDir, 'styles');
const targetStylesDir = path.join(assetsDir, 'styles');
fs.mkdirSync(targetStylesDir, { recursive: true });

// Copy styles.css with corrected font paths
let stylesContent = fs.readFileSync(path.join(stylesDir, 'styles.css'), 'utf8');
stylesContent = stylesContent.replace(/url\(['"]\.\.\/assets\/fonts\/WEB\//g, "url('../fonts/WEB/");
fs.writeFileSync(path.join(targetStylesDir, 'styles.css'), stylesContent);

fs.copyFileSync(path.join(stylesDir, 'pdfexport.css'), path.join(targetStylesDir, 'pdfexport.css'));

// Copy JS modules
const modulesDir = path.join(clientDir, 'modules');
const targetModulesDir = path.join(assetsDir, 'modules');
fs.mkdirSync(targetModulesDir, { recursive: true });

function copyDirectory(src, dest) {
    if (!fs.existsSync(dest)) {
        fs.mkdirSync(dest, { recursive: true });
    }
    
    const items = fs.readdirSync(src);
    for (const item of items) {
        const srcPath = path.join(src, item);
        const destPath = path.join(dest, item);
        
        if (fs.statSync(srcPath).isDirectory()) {
            copyDirectory(srcPath, destPath);
        } else {
            fs.copyFileSync(srcPath, destPath);
        }
    }
}

copyDirectory(modulesDir, targetModulesDir);

// Copy WEB fonts
const webFontsDir = path.join(__dirname, '..', 'WEB');
const targetFontsDir = path.join(assetsDir, 'fonts');
if (fs.existsSync(webFontsDir)) {
    console.log('📝 Copying WEB fonts...');
    copyDirectory(webFontsDir, targetFontsDir);
}

// Also copy other fonts if they exist
const srcFontsDir = path.join(__dirname, '..', 'src', 'assets', 'fonts');
if (fs.existsSync(srcFontsDir)) {
    // Copy additional fonts, but don't overwrite WEB fonts
    const items = fs.readdirSync(srcFontsDir);
    for (const item of items) {
        const srcPath = path.join(srcFontsDir, item);
        const destPath = path.join(targetFontsDir, item);
        
        if (!fs.existsSync(destPath)) {
            if (fs.statSync(srcPath).isDirectory()) {
                copyDirectory(srcPath, destPath);
            } else {
                fs.copyFileSync(srcPath, destPath);
            }
        }
    }
}

// Generate font CSS
const fontCSS = `
/* ABeZeh Font Family - Complete Set */
@font-face {
    font-family: 'ABeZeh';
    src: url('assets/fonts/WEB/ABeZeh-Regular.eot');
    src: url('assets/fonts/WEB/ABeZeh-Regular.eot?#iefix') format('embedded-opentype'),
         url('assets/fonts/WEB/ABeZeh-Regular.woff2') format('woff2'),
         url('assets/fonts/WEB/ABeZeh-Regular.woff') format('woff');
    font-weight: 400;
    font-style: normal;
    font-display: swap;
}

@font-face {
    font-family: 'ABeZeh';
    src: url('assets/fonts/WEB/ABeZeh-Italic.eot');
    src: url('assets/fonts/WEB/ABeZeh-Italic.eot?#iefix') format('embedded-opentype'),
         url('assets/fonts/WEB/ABeZeh-Italic.woff2') format('woff2'),
         url('assets/fonts/WEB/ABeZeh-Italic.woff') format('woff');
    font-weight: 400;
    font-style: italic;
    font-display: swap;
}

@font-face {
    font-family: 'ABeZeh';
    src: url('assets/fonts/WEB/ABeZeh-Thin.eot');
    src: url('assets/fonts/WEB/ABeZeh-Thin.eot?#iefix') format('embedded-opentype'),
         url('assets/fonts/WEB/ABeZeh-Thin.woff2') format('woff2'),
         url('assets/fonts/WEB/ABeZeh-Thin.woff') format('woff');
    font-weight: 100;
    font-style: normal;
    font-display: swap;
}

@font-face {
    font-family: 'ABeZeh';
    src: url('assets/fonts/WEB/ABeZeh-ThinItalic.eot');
    src: url('assets/fonts/WEB/ABeZeh-ThinItalic.eot?#iefix') format('embedded-opentype'),
         url('assets/fonts/WEB/ABeZeh-ThinItalic.woff2') format('woff2'),
         url('assets/fonts/WEB/ABeZeh-ThinItalic.woff') format('woff');
    font-weight: 100;
    font-style: italic;
    font-display: swap;
}

@font-face {
    font-family: 'ABeZeh';
    src: url('assets/fonts/WEB/ABeZeh-ExtraLight.eot');
    src: url('assets/fonts/WEB/ABeZeh-ExtraLight.eot?#iefix') format('embedded-opentype'),
         url('assets/fonts/WEB/ABeZeh-ExtraLight.woff2') format('woff2'),
         url('assets/fonts/WEB/ABeZeh-ExtraLight.woff') format('woff');
    font-weight: 200;
    font-style: normal;
    font-display: swap;
}

@font-face {
    font-family: 'ABeZeh';
    src: url('assets/fonts/WEB/ABeZeh-ExtraLightItalic.eot');
    src: url('assets/fonts/WEB/ABeZeh-ExtraLightItalic.eot?#iefix') format('embedded-opentype'),
         url('assets/fonts/WEB/ABeZeh-ExtraLightItalic.woff2') format('woff2'),
         url('assets/fonts/WEB/ABeZeh-ExtraLightItalic.woff') format('woff');
    font-weight: 200;
    font-style: italic;
    font-display: swap;
}

@font-face {
    font-family: 'ABeZeh';
    src: url('assets/fonts/WEB/ABeZeh-Medium.eot');
    src: url('assets/fonts/WEB/ABeZeh-Medium.eot?#iefix') format('embedded-opentype'),
         url('assets/fonts/WEB/ABeZeh-Medium.woff2') format('woff2'),
         url('assets/fonts/WEB/ABeZeh-Medium.woff') format('woff');
    font-weight: 500;
    font-style: normal;
    font-display: swap;
}

@font-face {
    font-family: 'ABeZeh';
    src: url('assets/fonts/WEB/ABeZeh-MediumItalic.eot');
    src: url('assets/fonts/WEB/ABeZeh-MediumItalic.eot?#iefix') format('embedded-opentype'),
         url('assets/fonts/WEB/ABeZeh-MediumItalic.woff2') format('woff2'),
         url('assets/fonts/WEB/ABeZeh-MediumItalic.woff') format('woff');
    font-weight: 500;
    font-style: italic;
    font-display: swap;
}

@font-face {
    font-family: 'ABeZeh';
    src: url('assets/fonts/WEB/ABeZeh-Bold.eot');
    src: url('assets/fonts/WEB/ABeZeh-Bold.eot?#iefix') format('embedded-opentype'),
         url('assets/fonts/WEB/ABeZeh-Bold.woff2') format('woff2'),
         url('assets/fonts/WEB/ABeZeh-Bold.woff') format('woff');
    font-weight: 700;
    font-style: normal;
    font-display: swap;
}

@font-face {
    font-family: 'ABeZeh';
    src: url('assets/fonts/WEB/ABeZeh-BoldItalic.eot');
    src: url('assets/fonts/WEB/ABeZeh-BoldItalic.eot?#iefix') format('embedded-opentype'),
         url('assets/fonts/WEB/ABeZeh-BoldItalic.woff2') format('woff2'),
         url('assets/fonts/WEB/ABeZeh-BoldItalic.woff') format('woff');
    font-weight: 700;
    font-style: italic;
    font-display: swap;
}

@font-face {
    font-family: 'ABeZeh';
    src: url('assets/fonts/WEB/ABeZeh-ExtraBold.eot');
    src: url('assets/fonts/WEB/ABeZeh-ExtraBold.eot?#iefix') format('embedded-opentype'),
         url('assets/fonts/WEB/ABeZeh-ExtraBold.woff2') format('woff2'),
         url('assets/fonts/WEB/ABeZeh-ExtraBold.woff') format('woff');
    font-weight: 800;
    font-style: normal;
    font-display: swap;
}

@font-face {
    font-family: 'ABeZeh';
    src: url('assets/fonts/WEB/ABeZeh-ExtraBoldItalic.eot');
    src: url('assets/fonts/WEB/ABeZeh-ExtraBoldItalic.eot?#iefix') format('embedded-opentype'),
         url('assets/fonts/WEB/ABeZeh-ExtraBoldItalic.woff2') format('woff2'),
         url('assets/fonts/WEB/ABeZeh-ExtraBoldItalic.woff') format('woff');
    font-weight: 800;
    font-style: italic;
    font-display: swap;
}

/* Apply ABeZeh font to everything */
body {
    font-family: 'ABeZeh', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif;
}

* {
    font-family: inherit;
}
`;

// Create static index.html with proper structure and fonts
const indexHtml = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Viewer - Static Version</title>
    <style>
${fs.readFileSync(path.join(targetStylesDir, 'styles.css'), 'utf8')}
    </style>
    <style>
        ${fontCSS}
        
        .upload-area {
            border: 2px dashed #ccc;
            border-radius: 8px;
            padding: 40px;
            text-align: center;
            margin: 20px;
            background: #f9f9f9;
        }
        
        .upload-area.drag-over {
            border-color: #34a3d7;
            background: #e8f4f8;
        }
        
        .welcome-message {
            text-align: center;
            padding: 40px;
            color: #666;
        }
    </style>
</head>
<body data-file-id="static-mode">
    <div class="container viewer-container">
        <header>
            <div class="header-controls">
                <div class="upload-controls">
                    <input type="file" id="file-input" accept=".xlsx,.xls" style="display: none;">
                    <button id="upload-button" class="control-button">📁 Upload Excel File</button>
                    <span id="file-info" style="margin-left: 10px; color: #666;"></span>
                </div>
                <div class="export-controls">
                    <select id="export-column">
                        <option value="">Select a column...</option>
                    </select>
                    <button id="preview-pdf" class="control-button">Preview PDF</button>
                    <button id="export-pdf" class="control-button">Export to PDF</button>
                </div>
            </div>
        </header>
        
        <main>
            <div class="toolbar">
                <button id="collapse-all" class="toolbar-button">Collapse All</button>
                <button id="expand-all" class="toolbar-button">Expand All</button>
                <label class="toolbar-checkbox">
                    <input type="checkbox" id="show-empty-cells" checked>
                    Show Empty Cells
                </label>
                
                <div class="control-group">
                    <div class="toggle-buttons" id="template-mode-toggle">
                        <button type="button" class="toggle-btn" data-value="hierarchy">Nested View</button>
                        <button type="button" class="toggle-btn active" data-value="list">List View</button>
                    </div>
                </div>
            </div>
            
            <div class="style-controls">
                <div class="control-group">
                    <label for="font-size">Font Size:</label>
                    <select id="font-size">
                        <option value="12px">Small</option>
                        <option value="14px" selected>Medium</option>
                        <option value="16px">Large</option>
                        <option value="18px">Extra Large</option>
                    </select>
                </div>
                
                <div class="control-group">
                    <label for="theme">Theme:</label>
                    <select id="theme">
                        <option value="default" selected>Default</option>
                        <option value="dark">Dark</option>
                        <option value="light">Light</option>
                        <option value="colorful">Colorful</option>
                    </select>
                </div>
                
                <div class="control-group">
                    <label for="cell-padding">Cell Padding:</label>
                    <select id="cell-padding">
                        <option value="2px">Compact</option>
                        <option value="5px" selected>Normal</option>
                        <option value="10px">Spacious</option>
                    </select>
                </div>
                
                <div class="control-group">
                    <label for="show-grid">Show Grid:</label>
                    <input type="checkbox" id="show-grid" checked>
                </div>
            </div>
            
            <div class="sheet-tabs" id="sheet-tabs">
                <!-- Sheet tabs will be added dynamically -->
            </div>
            
            <div class="sheet-content-container" id="sheet-content-container">
                <div class="welcome-message">
                    <h2>📊 Excel Viewer</h2>
                    <p>Upload an Excel file (.xlsx or .xls) to get started</p>
                    <div class="upload-area" id="upload-area">
                        <p>Click "Upload Excel File" above or drag & drop a file here</p>
                    </div>
                </div>
            </div>
        </main>
        
        <footer>
            <p>Excel Viewer &copy; ${new Date().getFullYear()}</p>
        </footer>
    </div>
    
    <!-- Global data setup -->
    <script>
        window.excelData = {
            fileId: "static-mode",
            sheetsLoaded: {},
            filename: "",
            sheets: {},
            activeSheetId: null
        };
    </script>
    
    <!-- External libraries -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <script src="assets/modules/html2canvas.min.js"></script>
    <script src="assets/modules/jspdf.umd.min.js"></script>
    
    <!-- App modules -->
    <script src="assets/modules/viewer/core.js"></script>
    <script src="assets/modules/viewer/export.js"></script>
    <script src="assets/modules/viewer/data-processing.js"></script>
    <script src="assets/modules/viewer/ui-controls.js"></script>
    <script src="assets/modules/viewer/node-management.js"></script>
    <script src="assets/modules/live-editor.js"></script>
    <script src="assets/modules/viewer/node-renderer.js"></script>
    <script src="assets/modules/viewer/sheet-management.js"></script>
    <script src="assets/modules/viewer/structure-builder.js"></script>
    <script src="assets/modules/viewer.js"></script>
    
    <!-- Static app initialization -->
    <script>
        // Debug script loading
        console.log('All scripts loaded, ExcelViewerCore:', window.ExcelViewerCore);
        
        document.addEventListener('DOMContentLoaded', function() {
            const fileInput = document.getElementById('file-input');
            const uploadButton = document.getElementById('upload-button');
            const fileInfo = document.getElementById('file-info');
            const uploadArea = document.getElementById('upload-area');
            
            // File upload handling
            uploadButton.addEventListener('click', () => {
                fileInput.click();
            });
            
            // Drag and drop
            uploadArea.addEventListener('dragover', (e) => {
                e.preventDefault();
                uploadArea.classList.add('drag-over');
            });
            
            uploadArea.addEventListener('dragleave', () => {
                uploadArea.classList.remove('drag-over');
            });
            
            uploadArea.addEventListener('drop', (e) => {
                e.preventDefault();
                uploadArea.classList.remove('drag-over');
                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    handleFile(files[0]);
                }
            });
            
            fileInput.addEventListener('change', function(e) {
                const file = e.target.files[0];
                if (file) {
                    handleFile(file);
                }
            });
            
            async function handleFile(file) {
                fileInfo.textContent = \`Loading \${file.name} (\${(file.size / 1024 / 1024).toFixed(2)} MB)...\`;
                
                try {
                    const arrayBuffer = await file.arrayBuffer();
                    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                    
                    // Setup sheet data with filtering (same as server)
                    const sheets = [];
                    workbook.SheetNames.forEach((sheetName, originalIndex) => {
                        const worksheet = workbook.Sheets[sheetName];
                        
                        // Skip hidden sheets - multiple approaches for XLSX.js
                        const sheetInfo = workbook.Workbook?.Sheets?.[originalIndex];
                        const isHidden = sheetInfo?.Hidden === 1 || sheetInfo?.Hidden === 2 || 
                                       sheetInfo?.state === 'hidden' || sheetInfo?.state === 'veryHidden' ||
                                       worksheet['!hidden'] || worksheet.Hidden ||
                                       // Also skip obvious non-data sheets
                                       sheetName.toLowerCase().includes('draaitabel') ||
                                       sheetName.toLowerCase().includes('pivot') ||
                                       (sheetName.startsWith('Blad') && sheetName.match(/^Blad\d+$/));
                        
                        if (isHidden) {
                            console.log(\`Skipping hidden/system sheet: \${sheetName}\`);
                            return;
                        }
                        
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                        
                        // Skip empty sheets (same as server filtering)
                        if (jsonData.length === 0 || (jsonData.length === 1 && jsonData[0].filter(Boolean).length === 0)) {
                            console.log(\`Skipping empty sheet: \${sheetName}\`);
                            return;
                        }
                        
                        sheets.push({
                            id: sheets.length + 1, // Sequential ID for non-filtered sheets
                            name: sheetName,
                            data: jsonData
                        });
                    });
                    
                    // EXACT same data structure as server version (viewer.ejs lines 119-122)
                    window.excelData = {
                        fileId: 'static-client-' + Date.now(),
                        sheetsLoaded: {} // All sheets will be loaded via JS (just like server version)
                    };
                    
                    // Mock the server API by intercepting fetch calls to /api/sheet/
                    const originalFetch = window.fetch;
                    window.fetch = function(url, options) {
                        if (url.includes('/api/sheet/')) {
                            // Extract sheet ID from URL: /api/sheet/fileId/sheetId
                            const sheetId = url.split('/').pop();
                            console.log('Mocking API call for sheet:', sheetId);
                            
                            // Return mocked response
                            const sheetData = sheets.find(s => s.id == sheetId);
                            const headers = (sheetData?.data[0] || []).map(h => String(h || ''));
                            
                            return Promise.resolve({
                                ok: true,
                                json: () => Promise.resolve({
                                    data: {
                                        headers: headers,
                                        data: sheetData?.data || [],
                                        isRawData: true,
                                        needsConfiguration: true,
                                        root: { type: 'root', children: [], level: -1 }
                                    }
                                })
                            });
                        }
                        // For other URLs, use original fetch
                        return originalFetch.apply(this, arguments);
                    };
                    
                    // Create sheet tabs
                    const sheetTabs = document.getElementById('sheet-tabs');
                    sheetTabs.innerHTML = '';
                    sheets.forEach((sheet, index) => {
                        const button = document.createElement('button');
                        button.className = \`tab-button \${index === 0 ? 'active' : ''}\`;
                        button.setAttribute('data-sheet-id', sheet.id);
                        button.textContent = sheet.name;
                        sheetTabs.appendChild(button);
                    });
                    sheetTabs.style.display = 'flex';
                    
                    // Create sheet content containers
                    const container = document.getElementById('sheet-content-container');
                    container.innerHTML = '';
                    sheets.forEach((sheet, index) => {
                        const div = document.createElement('div');
                        div.className = \`sheet-content \${index === 0 ? 'active' : ''}\`;
                        div.id = \`sheet-\${sheet.id}\`;
                        div.innerHTML = '<div class="loading-indicator"><p>Loading data...</p></div>';
                        container.appendChild(div);
                    });
                    
                    fileInfo.textContent = \`\${file.name} (\${sheets.length} sheets)\`;
                    
                    // EXACT copy of how the working server version initializes
                    setTimeout(() => {
                        // Initialize core module first (like viewer.js line 3)
                        window.ExcelViewerCore.initializeCore();
                        
                        // Get DOM elements from core module (like viewer.js lines 6-18)
                        const { tabButtons, sheetContents } = window.ExcelViewerCore.getDOMElements();
                        
                        // Add tab switching with lazy loading (EXACT copy from viewer.js lines 76-102)
                        tabButtons.forEach(button => {
                            button.addEventListener('click', async () => {
                                const sheetId = button.getAttribute('data-sheet-id');
                                window.excelData.activeSheetId = sheetId;
                                
                                // Update active tab
                                tabButtons.forEach(btn => btn.classList.remove('active'));
                                button.classList.add('active');
                                
                                // Show the selected sheet content
                                sheetContents.forEach(content => {
                                    if (content.id === \`sheet-\${sheetId}\`) {
                                        content.classList.add('active');
                                        
                                        // Check if we need to load this sheet
                                        if (window.excelData && !window.excelData.sheetsLoaded[sheetId]) {
                                            window.ExcelViewerSheetManager.loadSheetData(sheetId);
                                        } else {
                                            // Update export column dropdown with loaded data
                                            window.ExcelViewerCore.populateExportColumnDropdown();
                                        }
                                    } else {
                                        content.classList.remove('active');
                                    }
                                });
                            });
                        });
                        
                        // Initialize the application (like viewer.js line 105)
                        window.ExcelViewerUIControls.initialize();
                        
                        // Now trigger the first tab and ensure it's marked as active
                        setTimeout(() => {
                            // Reset all tabs and activate first one
                            tabButtons.forEach(btn => btn.classList.remove('active'));
                            const firstTab = tabButtons[0];
                            if (firstTab) {
                                firstTab.classList.add('active');
                                console.log('Triggering first tab with proper event handlers...');
                                firstTab.click();
                            }
                        }, 100);
                    }, 500);
                    
                } catch (error) {
                    console.error('Error processing Excel file:', error);
                    fileInfo.textContent = 'Error loading file: ' + error.message;
                    fileInfo.style.color = 'red';
                }
            }
        });
    </script>
</body>
</html>`;

fs.writeFileSync(path.join(distDir, 'index.html'), indexHtml);

// Create README for deployment
const readme = `# Excel Viewer - Static Build

This is a static build of the Excel Viewer application that can be deployed to any web hosting service.

## Features

- Client-side Excel file processing (no server required)
- PDF export functionality
- Custom ABeZeh font family integrated
- Hierarchical view and list template modes
- Supports .xlsx and .xls files
- Drag & drop file upload

## Deployment Instructions

1. Upload all contents of this folder to your web hosting service
2. Make sure index.html is accessible from your domain root
3. Ensure all assets are properly linked (check browser console for any 404 errors)

## Font Integration

The application uses the ABeZeh font family in all weights and styles:
- Thin (100) / ThinItalic
- ExtraLight (200) / ExtraLightItalic  
- Regular (400) / Italic
- Medium (500) / MediumItalic
- Bold (700) / BoldItalic
- ExtraBold (800) / ExtraBoldItalic

## Browser Requirements

- Modern browser with JavaScript enabled
- File API support for local file reading
- Canvas API for PDF generation
- Font loading support

## File Structure

- index.html - Main application file with embedded fonts
- assets/styles/ - CSS stylesheets
- assets/modules/ - JavaScript modules
- assets/fonts/ - ABeZeh font files (EOT, WOFF, WOFF2)
`;

fs.writeFileSync(path.join(distDir, 'README.md'), readme);

console.log('✅ Static build completed!');
console.log('📝 Integrated ABeZeh fonts from WEB folder');
console.log('📁 Files generated in: ./dist/');
console.log('🚀 Upload the contents of ./dist/ to your hosting service');
console.log('');
console.log('💡 Note: This version uses client-side Excel processing and works without a Node.js server'); 