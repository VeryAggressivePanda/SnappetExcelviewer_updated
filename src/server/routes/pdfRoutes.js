const express = require('express');
const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs').promises;
const router = express.Router();

// Ensure temp directory exists
const tempDir = path.join(__dirname, '../../temp');
fs.mkdir(tempDir, { recursive: true }).catch(console.error);

// Helper function to truncate text with middle ellipsis
function truncateMiddle(text, maxLength = 35) {
  if (!text || text.length <= maxLength) return text;
  
  const start = Math.floor(maxLength / 2);
  const end = Math.ceil(maxLength / 2);
  
  return text.substring(0, start) + '...' + text.substring(text.length - end);
}

// Function to generate node HTML
function renderNode(node, level) {
  // Make sure we're using the node's actual level rather than an incremented position
  // This ensures negative levels are properly represented
  const nodeLevel = node.level !== undefined ? node.level : level;
  
  // Handle negative levels by using absolute value for CSS class names
  // but preserve the actual level in data attributes
  const cssLevel = Math.abs(nodeLevel);
  
  // Start node HTML with level-specific classes
  let nodeHtml = `<div class="level-${cssLevel}-node${node.isEmpty ? ' level-' + cssLevel + '-empty' : ''}" 
    data-level="${nodeLevel}" 
    data-column-name="${node.columnName || ''}" 
    data-column-index="${node.columnIndex || ''}"
    data-is-empty="${node.isEmpty ? 'true' : 'false'}">`;
  
  // Check if node has children to determine if it's a parent
  const isParent = node.children && node.children.length > 0;
  
  // Add header - for parent nodes, includes both column name AND cell value
  nodeHtml += `<div class="level-${cssLevel}-header${node.isEmpty ? ' level-' + cssLevel + '-empty-header' : ''}">`;
  
  if (isParent) {
    // For parent nodes: only show title if node has a meaningful specific value
    const columnName = node.columnName || '';
    const nodeValue = node.value || '';
    
    // Check if this is a meaningful value (not just column name, "New Container", or empty)
    const isMeaningfulValue = nodeValue && 
                             nodeValue.trim() !== '' && 
                             nodeValue !== columnName && 
                             nodeValue !== 'New Container' &&
                             nodeValue !== 'Container';
    
    if (isMeaningfulValue) {
      // Node has a meaningful specific value - show just the value (like "Week 1", "Les 1", etc.)
      const truncatedValue = truncateMiddle(nodeValue);
      
      // Add title attribute for tooltip on hover
      const titleAttr = nodeValue !== truncatedValue ? 
        `title="${nodeValue}"` : '';
      
      nodeHtml += `<div class="level-${cssLevel}-title" ${titleAttr}>
        <span class="cell-value">${truncatedValue}</span>
      </div>`;
    } else {
      // Node has no meaningful value - hide the title but keep the header structure
      nodeHtml += `<div class="level-${cssLevel}-title" style="display: none;"></div>`;
    }
  } else {
    // For leaf nodes: show the actual value if it exists, otherwise show column name
    const nodeValue = node.value || '';
    const columnName = node.columnName || '';
    
    if (nodeValue && nodeValue.trim() !== '') {
      nodeHtml += `<div class="level-${cssLevel}-title">${nodeValue}</div>`;
    } else {
      nodeHtml += `<div class="level-${cssLevel}-title">${columnName}</div>`;
    }
  }
  
  nodeHtml += `</div>`; // Close header
  
  // Add content section
  nodeHtml += `<div class="level-${cssLevel}-content${node.isEmpty ? ' level-' + cssLevel + '-empty-content' : ''}">`;
  
  if (isParent) {
    // For parent nodes: content contains children
    if (node.children && node.children.length > 0) {
      // Determine layout based on child count and level
      let layoutClass = 'layout-horizontal';
      let childCountClass = '';
      
      if (node.children.length === 1) {
        childCountClass = 'layout-single-child';
      } else if (node.children.length === 2) {
        childCountClass = 'layout-two-children';
      } else if (node.children.length === 3) {
        childCountClass = 'layout-three-children';
      } else if (node.children.length > 3) {
        childCountClass = 'layout-many-children';
      }
      
      // Level 0 should typically use vertical layout for better PDF formatting
      if (cssLevel === 0) {
        layoutClass = 'layout-vertical';
        childCountClass = ''; // Reset for vertical layout
      }
      
      nodeHtml += `<div class="level-${cssLevel}-children ${layoutClass} ${childCountClass}" data-child-count="${node.children.length}">`;
      
      // Pass the node's actual level to children, not an incremented one
      node.children.forEach(child => {
        nodeHtml += renderNode(child, nodeLevel + 1);
      });
      
      nodeHtml += `</div>`; // Close children container
    }
  } else {
    // For leaf nodes: content contains cell value
    nodeHtml += node.value || '';
  }
  
  nodeHtml += `</div>`;
  
  // Process properties for all levels (non-hierarchy columns)
  if (node.properties && node.properties.length > 0) {
    const filteredProperties = node.properties.filter(prop => 
      prop.columnName !== "Column 9" && prop.columnName !== "Column 10"
    );
    
    if (filteredProperties.length > 0) {
      nodeHtml += `<div class="level-${cssLevel}-properties">`;
      
      filteredProperties.forEach(prop => {
        const isEmpty = !prop.value || prop.value.trim() === '';
        const propLevel = cssLevel + 1;
        
        // Create a property node
        nodeHtml += `<div class="level-${propLevel}-node${isEmpty ? ' level-' + propLevel + '-empty' : ''}" 
          data-column-name="${prop.columnName}" 
          data-column-index="${prop.columnIndex}">`;
        
        // Property header - contains ONLY column name
        nodeHtml += `<div class="level-${propLevel}-header${isEmpty ? ' level-' + propLevel + '-empty-header' : ''}">
          <div class="level-${propLevel}-title">${prop.columnName}</div>
        </div>`;
        
        // Property content - contains cell value
        nodeHtml += `<div class="level-${propLevel}-content${isEmpty ? ' level-' + propLevel + '-empty-content' : ''}">
          ${prop.value || ''}
        </div>`;
        
        nodeHtml += `</div>`; // Close property node
      });
      
      nodeHtml += `</div>`; // Close properties container
    }
  }
  
  nodeHtml += `</div>`; // Close node
  
  return nodeHtml;
}

// Function to generate PDF content
function generatePdfContent(data, title) {
  // Link to external PDF CSS file
  const styles = `
    <link rel="stylesheet" href="/styles/pdfexport.css">
    <style>
      /* Base styles */
      body {
        font-family: 'AbeZeh', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', sans-serif;
        line-height: 1.6;
        color: #333;
        margin: 0;
        padding: 0;
      }
      
      .pdf-container {
        width: 210mm;
        margin: 0 auto;
        padding: 0;
      }
      
      /* Basic node styles */
      .excel-node {
        margin-bottom: 1rem;
      }
      
      /* PDF title styles */
      .pdf-title {
        font-size: 1.4em !important;
        font-weight: bold !important;
        color: #34a3d7 !important;
        text-align: left !important;
        margin: 0 0 1rem 0 !important;
        padding: 0.75rem 1.5rem !important;
        background-color: rgba(52, 163, 215, 0.1) !important;
        border-radius: 25px !important;
        page-break-after: avoid !important;
        page-break-inside: avoid !important;
        display: block !important;
        width: fit-content !important;
        box-sizing: border-box !important;
      }
      
      /* Make sure h1.pdf-title specifically gets these styles */
      h1.pdf-title {
        font-size: 1.4em !important;
        font-weight: bold !important;
        color: #34a3d7 !important;
        margin: 0 0 1rem 0 !important;
        padding: 0.75rem 1.5rem !important;
        background-color: rgba(52, 163, 215, 0.1) !important;
        border-radius: 25px !important;
        page-break-after: avoid !important;
        page-break-inside: avoid !important;
        display: block !important;
        width: fit-content !important;
        box-sizing: border-box !important;
      }
      
      /* New wrapper container to keep title and content together */
      .pdf-title-and-content {
        page-break-inside: avoid !important;
        page-break-after: auto !important;
        display: block !important;
        width: 100% !important;
      }
      
      /* Ensure pdf-content section flows correctly after title */
      .pdf-content {
        page-break-before: avoid !important;
        display: block !important;
      }
      
      /* Make the first content element stick with title */
      .pdf-section {
        page-break-before: avoid !important;
        display: block !important;
      }
      
      .pdf-section {
        margin: 0;
        padding: 0;
      }
      
      /* Common styles for all level nodes */
      [class*="level-"] {
        box-sizing: border-box;
      }
      
      /* Level node base styles */
      [class*="-node"] {
        display: flex;
        flex-direction: column;
        margin: 0.25rem 0;
        border-radius: 4px;
        overflow: hidden;
        width: 100%;
      }
      
      /* Level headers base styles */
      [class*="-header"] {
        display: flex;
        align-items: center;
        padding: 0.5rem;
        background-color: rgba(52, 163, 215, 0.05);
        width: 100%;
      }
      
      /* Level content base styles */
      [class*="-content"] {
        word-break: break-word;
      }
      
      /* Level titles base styles */
      [class*="-title"] {
        font-weight: 500;
        flex: 1;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        max-width: 100%;
      }
      
      /* Cell value and column name styles */
      .column-name {
        font-weight: 600;
        color: #166d9e;
        margin-right: 0.25em;
      }
      
      .cell-value {
        font-weight: 500;
        color: #333;
      }
      
      /* Level-0 (Course) styling */
      .level-0-node {
        background-color: #ffffff;
        border-radius: 8px;
        box-shadow: 0 3px 7px rgba(51, 51, 51, 0.2);
        margin: .5rem 0;
      }
      
      /* Only apply page break avoidance to level-0 nodes that are NOT the first one */
      .level-0-node:not(:first-of-type) {
        page-break-inside: avoid;
      }
      
      /* The first level-0 node should flow with the title */
      .level-0-node:first-of-type {
        page-break-before: avoid !important;
      }
      
      .level-0-header {
        background-color: white;
        border: none;
        padding: 1rem;
      }
      
      .level-0-title {
        font-size: 1.2rem;
        font-weight: bold;
        color: #34a3d7;
      }
      
      .level-0-content {
        padding: 1rem;
        color: #333;
        font-size: 1rem;
      }
      
      /* Set padding to 0 when level-0-node is collapsed */
      .level-0-collapsed .level-0-content {
        padding: 0;
      }
      
      .level-0-children {
        display: flex;
        flex-direction: column;
        width: 100%;
        padding: 0;
        gap: 1rem;
      }
      
      /* Level 1 styles */
      .level-1-node {
        background-color: transparent !important;
        border-bottom: 1px solid rgba(52, 163, 215, 0.2);
        padding-bottom: 1rem;
        margin-top: 0;
        width: 100%;
      }
      
      .level-1-header {
        background-color: transparent !important;
        padding: 0;
      }
      
      .level-1-title {
        font-size: 1rem;
        font-weight: bold;
        color: #333;
      }
      
      .level-1-content {
        padding: 0;
        color: #333;
      }
      
      .level-1-children {
        display: flex;
        flex-direction: row;
        flex-wrap: wrap;
        width: 100%;
        padding: 0;
        margin-left: 0;
        gap: 0.5rem;
      }
      
      /* Level 2 styles */
      .level-2-node {
        display: flex;
        flex-direction: column;
        background-color: rgb(52 163 215 / 5%);
        border-radius: 4px;
        margin: 0.5rem 0;
        padding: 0.5rem;
        width: 100%;
        page-break-inside: avoid;
      }
      
      .level-2-header {
        background-color: rgba(74, 111, 165, 0.05);
        border-radius: 4px 4px 0 0;
        padding: 0.75rem;
      }
      
      .level-2-title {
        font-size: 0.95rem;
        font-weight: 500;
        color: #333;
      }
      
      .level-2-content {
        padding: 0;
        color: #333;
        font-size: 0.95rem;
      }
      
      .level-2-children {
        display: flex;
        flex-direction: row;
        flex-wrap: wrap;
        width: 100%;
        gap: 0.5rem;
        padding: 0.5rem 0;
      }
      
      /* Level 3 styles */
      .level-3-node {
        display: flex;
        flex-direction: column;
        background-color: white;
        border-radius: 4px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        margin: 0;
        
       
      }
      
      .level-3-header {
        background-color: rgba(52, 163, 215, 0.1);
        padding: 0.5rem;
        border-bottom: 1px solid rgba(52, 163, 215, 0.1);
      }
      
      .level-3-title {
        font-size: 0.9rem;
        font-weight: 500;
        color: #166d9e;
      }
      
      .level-3-content {
        padding: 0.5rem;
        color: #333;
        font-size: 0.9rem;
      }
      
      /* Properties styling */
      .level-0-properties,
      .level-1-properties,
      .level-2-properties {
        display: flex;
        flex-direction: row;
        flex-wrap: wrap;
        width: 100%;
        gap: 0.5rem;
        padding: 0.5rem;
      }
      
      /* Empty node styling */
      .level-0-empty,
      .level-1-empty,
      .level-2-empty,
      .level-3-empty {
        border: 1px dashed rgba(52, 163, 215, 0.3);
        background-color: rgba(249, 249, 249, 0.7) !important;
      }
      
      .level-0-empty-header,
      .level-1-empty-header,
      .level-2-empty-header,
      .level-3-empty-header {
        color: #999;
        background-color: rgba(249, 249, 249, 0.9) !important;
      }
      
      .level-0-empty-content,
      .level-1-empty-content,
      .level-2-empty-content,
      .level-3-empty-content {
        color: #999;
        font-style: italic;
      }
      
      /* Layout classes */
      .level-0-vertical-layout {
        display: flex;
        flex-direction: column;
      }
      
      .level-1-grid-layout {
        display: flex;
        flex-direction: row;
        flex-wrap: wrap;
      }
      
      /* Child count-based layouts */
      .level-0-child-count-1 > .level-1-node {
        flex-basis: 100%;
      }
      
      .level-1-child-count-1 > .level-2-node {
        flex-basis: 100%;
      }
      
      /* Property styling based on data attributes */
      .level-0-property[data-column-name="Bron"],
      .level-1-property[data-column-name="Bron"],
      .level-2-property[data-column-name="Bron"],
      .level-0-property[data-column-name="Source"],
      .level-1-property[data-column-name="Source"],
      .level-2-property[data-column-name="Source"] {
        border-left: 3px solid #4a6fa5;
      }
      
      .level-0-property[data-column-name="Materiaal"],
      .level-1-property[data-column-name="Materiaal"],
      .level-2-property[data-column-name="Materiaal"],
      .level-0-property[data-column-name="Material"],
      .level-1-property[data-column-name="Material"],
      .level-2-property[data-column-name="Material"] {
        border-left: 3px solid #28a745;
      }
      
      .level-0-property[data-column-name="Type"],
      .level-1-property[data-column-name="Type"],
      .level-2-property[data-column-name="Type"] {
        border-left: 3px solid #dc3545;
      }
      
      .level-0-property[data-column-name="URL"],
      .level-1-property[data-column-name="URL"],
      .level-2-property[data-column-name="URL"],
      .level-0-property[data-column-name="Link"],
      .level-1-property[data-column-name="Link"],
      .level-2-property[data-column-name="Link"] {
        border-left: 3px solid #fd7e14;
      }
      
      .level-0-property[data-column-name="Beschrijving"],
      .level-1-property[data-column-name="Beschrijving"],
      .level-2-property[data-column-name="Beschrijving"],
      .level-0-property[data-column-name="Description"],
      .level-1-property[data-column-name="Description"],
      .level-2-property[data-column-name="Description"] {
        border-left: 3px solid #6610f2;
      }
      
      /* Empty nodes and properties */
      .level-0-empty,
      .level-1-empty,
      .level-2-empty {
        border: 1px dashed rgba(52, 163, 215, 0.3);
        background-color: rgba(249, 249, 249, 0.7) !important;
      }
      
      .level-0-empty-header,
      .level-1-empty-header,
      .level-2-empty-header {
        color: #999;
        background-color: rgba(249, 249, 249, 0.9) !important;
      }
      
      .level-0-empty-title,
      .level-1-empty-title,
      .level-2-empty-title {
        color: #999;
        font-style: italic;
      }
      
      .level-0-empty-property,
      .level-1-empty-property,
      .level-2-empty-property {
        border: 1px dashed rgba(204, 204, 204, 0.5);
        opacity: 0.8;
      }
      
      .level-0-empty-value,
      .level-1-empty-value,
      .level-2-empty-value {
        color: #999;
      }
      
      .empty-indicator {
        font-size: 0.8em;
        color: #999;
        background-color: #f0f0f0;
        padding: 0.1rem 0.3rem;
        border-radius: 0.2rem;
        margin-right: 0.3rem;
        font-weight: normal;
        font-style: normal;
      }
      
      .empty-value {
        font-size: 0.8em;
        color: #999;
        font-style: italic;
      }
      
      /* Print-specific styles */
      @media print {
        @page {
          size: A4;
          margin: 1cm;
        }
        
        body {
          width: 210mm;
        }
        
        .pdf-container {
          width: 100%;
        }
        
        * {
          -webkit-print-color-adjust: exact !important;
          print-color-adjust: exact !important;
        }
        
        /* Adjust level-3 nodes to be more responsive in print */
        .level-3-node {
          min-width: 100px;
          flex: 1 1 calc(20% - 0.5rem);
        }
        
        /* For very narrow content, stack elements vertically */
        @media (max-width: 480px) {
          .level-1-children {
            flex-direction: column;
          }
          
          .level-2-children {
            flex-direction: column;
          }
          
          .level-3-node {
            max-width: 100%;
            flex-basis: 100%;
          }
        }
      }
    </style>
  `;
  
  // Start HTML document
  let html = `
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>${title}</title>
      ${styles}
    </head>
    <body>
      <div class="pdf-container">
        <div class="pdf-title-and-content">
          <h1 class="pdf-title">${title}</h1>
          <div class="pdf-content">
            <div class="pdf-section">
  `;
  
  // Render the data as hierarchical nodes
  data.forEach(node => {
    html += renderNode(node, 0);
  });
  
  // Close HTML
  html += `
            </div>
          </div>
        </div>
      </div>
    </body>
    </html>
  `;
  
  return html;
}

// PDF generation endpoint
router.post('/generate-pdf', async (req, res) => {
  try {
    const { html, filename } = req.body;
    
    if (!html || !filename) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    // Create a unique filename
    const pdfFilename = `${filename.replace(/[^a-z0-9]/gi, '_')}_${Date.now()}.pdf`;
    const outputPath = path.join(tempDir, pdfFilename);
    
    // Set up browser for PDF generation
    const browser = await puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    
    const page = await browser.newPage();
    
    // Parse the data and generate HTML with styles
    const data = JSON.parse(html);
    const fullHtml = generatePdfContent(data, filename);
    
    await page.setContent(fullHtml, { waitUntil: 'networkidle0' });
    
    // Configure PDF options
    const pdfOptions = {
      path: outputPath,
      format: 'A4',
      printBackground: true,
      margin: {
        top: '1cm',
        right: '1cm',
        bottom: '1cm',
        left: '1cm'
      }
    };
    
    // Generate PDF
    await page.pdf(pdfOptions);
    await browser.close();
    
    // Send the PDF file
    res.sendFile(outputPath, {}, err => {
      if (err) {
        console.error('Error sending file:', err);
        res.status(500).json({ error: 'Could not send PDF file' });
      }
      
      // Delete the file after sending
      fs.unlink(outputPath, err => {
        if (err) console.error('Error deleting temp file:', err);
      });
    });
  } catch (error) {
    console.error('PDF generation error:', error);
    res.status(500).json({ error: 'Failed to generate PDF' });
  }
});

// HTML preview endpoint - returns exact same HTML that would be used for PDF
router.post('/preview-html', (req, res) => {
  try {
    const { html, filename } = req.body;
    
    if (!html || !filename) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    // Parse the data and generate HTML with styles - exactly the same as PDF
    const data = JSON.parse(html);
    const fullHtml = generatePdfContent(data, filename);
    
    // Send the HTML
    res.send(fullHtml);
  } catch (error) {
    console.error('HTML preview generation error:', error);
    res.status(500).json({ error: 'Failed to generate HTML preview' });
  }
});

module.exports = router; 