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
            @page {
              size: A4;
              margin: 0;
            }
            body {
              margin: 0;
              padding: 0;
              background: white;
              width: 210mm;
              height: 297mm;
            }
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
        top: '25mm',
        right: '25mm',
        bottom: '25mm',
        left: '25mm'
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