const express = require('express');
const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs').promises;
const router = express.Router();

// Ensure temp directory exists
const tempDir = path.join(__dirname, '..', 'temp');
fs.mkdir(tempDir, { recursive: true }).catch(console.error);

// PDF generation endpoint
router.post('/generate', async (req, res) => {
  const { html, filename } = req.body;
  
  if (!html || !filename) {
    return res.status(400).json({ error: 'HTML content and filename are required' });
  }

  const htmlPath = path.join(tempDir, 'export.html');
  const pdfPath = path.join(tempDir, 'export.pdf');

  try {
    // Write HTML to temporary file
    await fs.writeFile(htmlPath, html);

    // Launch browser
    const browser = await puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    // Create new page
    const page = await browser.newPage();

    // Set viewport to A4 size
    await page.setViewport({
      width: 794, // A4 width in pixels at 96 DPI
      height: 1123 // A4 height in pixels at 96 DPI
    });

    // Read the CSS file
    const cssPath = path.join(__dirname, '..', 'public', 'css', 'styles.css');
    const cssContent = await fs.readFile(cssPath, 'utf8');

    // Create HTML with styles
    const fullHtml = `
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="UTF-8">
          <style>
            ${cssContent}
            
            /* EXTREMELY SPECIFIC OVERRIDES */
            /* First, clear ALL backgrounds */
            div, header, section, body, html, * {
              background-color: transparent !important;
            }
            
            /* Then ONLY add background to level-2 node headers using a very specific selector */
            div.excel-node.level-2 > div.excel-node-header {
              background-color: rgba(74, 111, 165, 0.05) !important;
              -webkit-print-color-adjust: exact !important;
              print-color-adjust: exact !important;
            }
            
            /* Ensuring level-1 has absolutely no background */
            div.excel-node.level-1, 
            div.excel-node.level-1 > div.excel-node-header {
              background-color: transparent !important;
              background: none !important;
            }
            
            @page {
              size: A4;
            }
            body {
              font-family: 'AbeZeh', sans-serif;
              margin: 0;
              padding: 0;
              background: white;
              width: 210mm;
              height: 297mm;
              font-size: 10pt;
              line-height: 1.5;
              color: #333;
            }
            
            /* Blok styling (level-0) */
            .level-0 {
              box-shadow: 0 3px 7px rgba(51,51,51,0.2);
              margin: 1rem;
              background-color: transparent !important;
            }
            .level-0 > .excel-node-header {
              border: none;
              background-color: transparent !important;
            }
            .level-0 .node-title {
              font-size: 1.2rem;
              font-weight: bold;
              color: #34a3d7;
            }
            
            /* Week (level-1) - NO BACKGROUND */
            .excel-node.level-1 {
              background-color: transparent !important;
              border-bottom: 1px solid rgba(52, 163, 215, 0.2);
              padding-bottom: 1rem;
              margin-top: 0;
            }
            .excel-node.level-1 > .excel-node-header {
              background-color: transparent !important;
              box-shadow: none;
            }
            .level-1 .node-title {
              font-size: 1rem;
              font-weight: bold;
              color: #333;
            }
            
            /* Les (level-2) - WITH BACKGROUND */
            .excel-node.level-2 {
              break-inside: avoid;
              page-break-inside: avoid;
              margin: 0.5rem 0;
              padding: 0.5rem;
              background-color: transparent !important;
              border-radius: 4px;
              border: none;
            }
            
            /* PDF specific elements */
            .pdf-title {
              font-size: 1.2rem;
              font-weight: bold;
              color: #34a3d7;
              break-after: avoid;
              page-break-after: avoid;
              background-color: transparent !important;
            }
            
            /* Child elements */
            .excel-node-children {
              padding: 0;
              margin-left: 1rem;
            }
            
            /* Property styling */
            .property-row {
              background-color: rgb(52 163 215 / 3%) !important;
              border: none;
            }
            .excel-node-property {
              background-color: white !important;
              border: none;
            }
            
            /* Container styling */
            .pdf-container {
              transform-origin: top left;
              transform: scale(1);
            }
          </style>
        </head>
        <body>
          ${html}
        </body>
      </html>
    `;

    // Set content with styles
    await page.setContent(fullHtml, {
      waitUntil: 'networkidle0'
    });

    // Generate PDF
    await page.pdf({
      path: pdfPath,
      format: 'A4',
      printBackground: true,
      margin: {
        top: '15mm',
        right: '20mm',
        bottom: '15mm',
        left: '20mm'
      },
      preferCSSPageSize: true
    });

    // Close browser
    await browser.close();

    // Send PDF file
    res.download(pdfPath, filename, async (err) => {
      if (err) {
        console.error('Error sending file:', err);
      }
      
      // Clean up temporary files
      try {
        await fs.unlink(htmlPath);
        await fs.unlink(pdfPath);
      } catch (error) {
        console.error('Error cleaning up temporary files:', error);
      }
    });

  } catch (error) {
    console.error('Error generating PDF:', error);
    
    // Clean up temporary files in case of error
    try {
      await fs.unlink(htmlPath).catch(() => {});
      await fs.unlink(pdfPath).catch(() => {});
    } catch (cleanupError) {
      console.error('Error cleaning up temporary files:', cleanupError);
    }
    
    res.status(500).json({ error: 'Failed to generate PDF', details: error.message });
  }
});

module.exports = router; 