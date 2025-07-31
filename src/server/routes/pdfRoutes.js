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

// Function to intelligently determine if a node is really empty (same as frontend)
function determineIfNodeIsEmpty(node) {
  // If explicitly marked as empty, respect that
  if (node.isEmpty === true) {
    return true;
  }
  
  // Check if the node has meaningful content
  const hasValue = node.value && node.value.trim() !== '';
  const hasChildren = node.children && node.children.length > 0;
  const hasNonEmptyProperties = node.properties && 
    node.properties.some(prop => prop.value && prop.value.trim() !== '');
  
  // A node is empty if:
  // 1. It has no value AND no children AND no non-empty properties
  // 2. OR it only has a column name as value (like "Werkblad") but no actual content
  if (!hasValue && !hasChildren && !hasNonEmptyProperties) {
    return true;
  }
  
  // Special case: if the node's value is just the column name and has no other content
  // This catches cases like "Werkblad" nodes with empty content
  if (hasValue && node.columnName && 
      node.value.trim() === node.columnName.trim() && 
      !hasChildren && !hasNonEmptyProperties) {
    return true;
  }
  
  // If it's a parent node, check if all children are empty
  if (hasChildren) {
    const allChildrenEmpty = node.children.every(child => determineIfNodeIsEmpty(child));
    if (allChildrenEmpty && !hasValue && !hasNonEmptyProperties) {
      return true;
    }
  }
  
  return false;
}

// Function to render a single node to HTML string
function renderNode(node, level) {
  const nodeLevel = level;
  const cssLevel = level < 0 ? `_${Math.abs(level)}` : level;
  
  // Determine if this node should be considered empty
  const isReallyEmpty = determineIfNodeIsEmpty(node);
  
  // Start node HTML with level-specific classes
  let nodeHtml = `<div class="level-${cssLevel}-node${isReallyEmpty ? ' level-' + cssLevel + '-empty' : ''}" 
    data-level="${nodeLevel}" 
    data-column-name="${node.columnName || ''}" 
    data-column-index="${node.columnIndex || ''}"
    data-is-empty="${isReallyEmpty ? 'true' : 'false'}">`;
  
  // Check if node has children to determine if it's a parent
  const isParent = node.children && node.children.length > 0;
  
  // Skip header for level 0 nodes in PDF export (removes "Blok 1" etc.)
  if (cssLevel !== 0) {
    // Add header - for parent nodes, includes both column name AND cell value
    nodeHtml += `<div class="level-${cssLevel}-header${isReallyEmpty ? ' level-' + cssLevel + '-empty-header' : ''}">`;
    
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
  }

  // Add content section
  nodeHtml += `<div class="level-${cssLevel}-content${isReallyEmpty ? ' level-' + cssLevel + '-empty-content' : ''}">`;
  
  if (isParent) {
    // For parent nodes: content contains children
    if (node.children && node.children.length > 0) {
      // Determine layout based on node's layoutMode property (like in normal viewer)
      let layoutClass = 'layout-horizontal'; // Default
      let childCountClass = '';
      
      // Use node's layoutMode if it exists, otherwise use default logic
      if (node.layoutMode) {
        layoutClass = `layout-${node.layoutMode}`;
      } else {
        // Default behavior: use horizontal layout unless level 0 with many children
        layoutClass = 'layout-horizontal';
        
        // Only apply vertical layout to level 0 if there are many children (> 3)
        if (cssLevel === 0 && node.children.length > 3) {
          layoutClass = 'layout-vertical';
        }
      }
      
      // Set child count classes for horizontal layouts only
      if (layoutClass === 'layout-horizontal') {
        if (node.children.length === 1) {
          childCountClass = 'layout-single-child';
        } else if (node.children.length === 2) {
          childCountClass = 'layout-two-children';
        } else if (node.children.length === 3) {
          childCountClass = 'layout-three-children';
        } else if (node.children.length > 3) {
          childCountClass = 'layout-many-children';
        }
      }
      
      const fullLayoutClass = `${layoutClass}${childCountClass ? ' ' + childCountClass : ''}`;
      
      nodeHtml += `<div class="level-${cssLevel}-children ${fullLayoutClass}">`;
      
      // Render each child
      node.children.forEach(child => {
        nodeHtml += renderNode(child, level + 1);
      });
      
      nodeHtml += `</div>`; // Close children container
    }
  } else {
    // For leaf nodes: content might contain the cell value
    if (node.value && node.value.trim() !== '') {
      nodeHtml += node.value;
    }
  }
  
  nodeHtml += `</div>`; // Close content
  
  // Handle properties if they exist
  if (node.properties && node.properties.length > 0) {
    nodeHtml += `<div class="level-${cssLevel}-properties">`;
    
    node.properties.forEach(prop => {
      const propLevel = prop.level !== undefined ? prop.level : nodeLevel + 1;
      const propCssLevel = propLevel < 0 ? `_${Math.abs(propLevel)}` : propLevel;
      const isEmpty = !prop.value || prop.value.trim() === '' || prop.value === prop.columnName;
      
      // Property as individual node
      nodeHtml += `<div class="level-${propCssLevel}-node excel-node-property${isEmpty ? ' level-' + propCssLevel + '-empty' : ''}" 
        data-level="${propLevel}" 
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
  
  nodeHtml += `</div>`; // Close node
  
  return nodeHtml;
}

// Function to generate PDF content
function generatePdfContent(data, title, showEmptyCells = true) {
  // Read CSS file and inline it for PDF generation
  const path = require('path');
  const fs = require('fs');
  
  let pdfCss = '';
  try {
    const cssPath = path.join(__dirname, '../../client/styles/pdfexport.css');
    pdfCss = fs.readFileSync(cssPath, 'utf8');
  } catch (error) {
    console.error('Error reading pdfexport.css:', error);
  }
  
  const styles = `
    <style>
      /* Base font family */
      body {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', sans-serif;
        margin: 0;
        padding: 0;
      }
      
      /* Inlined PDF styles */
      ${pdfCss}
      
      /* Hide empty cells when hide-empty-cells class is applied to body */
      .hide-empty-cells .level-0-node.level-0-empty,
      .hide-empty-cells .level-1-node.level-1-empty,
      .hide-empty-cells .level-2-node.level-2-empty,
      .hide-empty-cells .level-3-node.level-3-empty,
      .hide-empty-cells .level-4-node.level-4-empty,
      .hide-empty-cells .level-_1-node.level-_1-empty,
      .hide-empty-cells .level-_2-node.level-_2-empty {
        display: none !important;
      }
      
      /* More specific: hide nodes that have both level class AND empty class */
      .hide-empty-cells [class*="level-"][class*="-node"][class*="-empty"] {
        display: none !important;
      }
      
      /* Hide empty content and headers specifically */
      .hide-empty-cells [class*="level-"][class*="-empty-content"],
      .hide-empty-cells [class*="level-"][class*="-empty-header"] {
        display: none !important;
      }
      
      /* Catch-all for any element with empty in class name */
      .hide-empty-cells [class*="-empty"] {
        display: none !important;
      }
    </style>
  `;
  
  // Determine body class based on showEmptyCells setting
  const bodyClass = showEmptyCells ? '' : ' class="hide-empty-cells"';
  

  
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
    <body${bodyClass}>
      <div class="pdf-container">
        <div class="pdf-title-and-content">
          <h1 class="pdf-title">Snappet Rekenen</h1>
          <h2 class="pdf-subtitle">${title}</h2>
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

// NEW: Function to generate PDF content from direct HTML (for list templates)
function generateListPdfContent(listHtml, title, showEmptyCells = true) {
  // Read CSS file and inline it for PDF generation
  const path = require('path');
  const fs = require('fs');
  
  let pdfCss = '';
  let listTemplateCss = '';
  try {
    const cssPath = path.join(__dirname, '../../client/styles/pdfexport.css');
    pdfCss = fs.readFileSync(cssPath, 'utf8');
    
    // Also include main styles.css for list template specific styles
    const mainCssPath = path.join(__dirname, '../../client/styles/styles.css');
    const mainCss = fs.readFileSync(mainCssPath, 'utf8');
    
    // Extract only list template related styles from main CSS
    listTemplateCss = extractListTemplateStyles(mainCss);
  } catch (error) {
    console.error('Error reading CSS files:', error);
  }
  
const styles = `
    <style>
      /* Base font family */
      body {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', sans-serif;
        margin: 0;
        padding: 0;
      }
      
      /* PDF export styles */
      ${pdfCss}
      
      /* List template specific styles */
      ${listTemplateCss}
      
      /* PDF-specific adjustments for list template */
      .sheet-container[data-is-list-template="true"] {
        width: 100% !important;
        max-width: 100% !important;
        margin: 0 !important;
        height: auto !important;
      }
      
      /* Ensure proper column layout for PDF */
      .sheet-container[data-is-list-template="true"]:not([data-has-courses="true"]) {
        display: flex !important;
        flex-direction: row !important;
        gap: 1rem !important;
      }
      
      .sheet-container[data-is-list-template="true"]:not([data-has-courses="true"]) > .level-0-node {
        flex: 1 !important;
        width: calc(33.333% - 0.67rem) !important;
        max-width: calc(33.333% - 0.67rem) !important;
        height: auto !important;
      }
      
      /* Remove fixed heights for PDF */
      .sheet-container[data-is-list-template="true"] .level-0-node > .level-0-content {
        height: auto !important;
        max-height: none !important;
        overflow: visible !important;
      }
      
      /* Hide empty cells when hide-empty-cells class is applied to body */
      .hide-empty-cells .level-0-node.level-0-empty,
      .hide-empty-cells .level-1-node.level-1-empty,
      .hide-empty-cells .level-2-node.level-2-empty,
      .hide-empty-cells .level-3-node.level-3-empty,
      .hide-empty-cells .level-4-node.level-4-empty,
      .hide-empty-cells .level-_1-node.level-_1-empty,
      .hide-empty-cells .level-_2-node.level-_2-empty {
        display: none !important;
      }
      
      /* More specific: hide nodes that have both level class AND empty class */
      .hide-empty-cells [class*="level-"][class*="-node"][class*="-empty"] {
        display: none !important;
      }
      
      /* Hide empty content and headers specifically */
      .hide-empty-cells [class*="level-"][class*="-empty-content"],
      .hide-empty-cells [class*="level-"][class*="-empty-header"] {
        display: none !important;
      }
      
      /* Catch-all for any element with empty in class name */
      .hide-empty-cells [class*="-empty"] {
        display: none !important;
      }
    </style>
`;
  
  // Determine body class based on showEmptyCells setting
  const bodyClass = showEmptyCells ? '' : ' class="hide-empty-cells"';
  
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
    <body${bodyClass}>
      <div class="pdf-container">
        <div class="pdf-title-and-content">
          <h1 class="pdf-title">Snappet Rekenen</h1>
          <h2 class="pdf-subtitle">${title}</h2>
          <div class="pdf-content">
            ${listHtml}
          </div>
        </div>
      </div>
    </body>
    </html>
`;
  
  return html;
}

// Function to extract list template specific styles from main CSS
function extractListTemplateStyles(mainCss) {
  // Extract styles that contain list template selectors
  const listTemplateSelectors = [
    '.sheet-container[data-is-list-template="true"]',
    '.level-0-node',
    '.level-1-node', 
    '.level-_1-node',
    '.level-0-header',
    '.level-1-header',
    '.level-_1-header',
    '.level-0-content',
    '.level-1-content',
    '.level-_1-content',
    '.level-0-children',
    '.level-1-children',
    '.level-_1-children'
  ];
  
  let extractedStyles = '';
  const lines = mainCss.split('\n');
  let inRelevantRule = false;
  let braceLevel = 0;
  let currentRule = '';
  
  for (const line of lines) {
    if (!inRelevantRule) {
      // Check if this line starts a relevant rule
      const isRelevant = listTemplateSelectors.some(selector => 
        line.includes(selector) || line.includes('[data-is-list-template')
      );
      
      if (isRelevant) {
        inRelevantRule = true;
        currentRule = line + '\n';
        braceLevel = (line.match(/{/g) || []).length - (line.match(/}/g) || []).length;
        
        if (braceLevel === 0) {
          extractedStyles += currentRule;
          currentRule = '';
          inRelevantRule = false;
        }
      }
    } else {
      currentRule += line + '\n';
      braceLevel += (line.match(/{/g) || []).length - (line.match(/}/g) || []).length;
      
      if (braceLevel === 0) {
        extractedStyles += currentRule;
        currentRule = '';
        inRelevantRule = false;
      }
    }
  }
  
  return extractedStyles;
}

// PDF generation endpoint
router.post('/generate-pdf', async (req, res) => {
  try {
    const { html, filename, showEmptyCells = true } = req.body;
    

    
    if (!html || !filename) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    // Create a unique filename
    const pdfFilename = `${filename.replace(/[^a-z0-9]/gi, '_')}_${Date.now()}.pdf`;
    const outputPath = path.join(tempDir, pdfFilename);
    
    // Set up browser for PDF generation
    const browser = await puppeteer.launch({
      headless: 'new',
      args: [
        '--no-sandbox', 
        '--disable-setuid-sandbox',
        '--disable-web-security',
        '--font-render-hinting=none'
      ]
    });
    
    const page = await browser.newPage();
    
    // Parse the data and generate HTML with styles
    const data = JSON.parse(html);

    const fullHtml = generatePdfContent(data, filename, showEmptyCells);
    
    await page.setContent(fullHtml, { waitUntil: 'networkidle0' });
    
    // Emulate screen media instead of print to match preview
    await page.emulateMediaType('screen');
    
    // Configure PDF options
    const pdfOptions = {
      path: outputPath,
      format: 'A4',
      printBackground: true,
      preferCSSPageSize: false,
      margin: {
        top: '20mm',
        right: '20mm', 
        bottom: '20mm',
        left: '20mm'
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
    const { html, filename, showEmptyCells = true } = req.body;
    

    
    if (!html || !filename) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    // Parse the data and generate HTML with styles - exactly the same as PDF
    const data = JSON.parse(html);

    const fullHtml = generatePdfContent(data, filename, showEmptyCells);
    
    // Send the HTML
    res.send(fullHtml);
  } catch (error) {
    console.error('HTML preview generation error:', error);
    res.status(500).json({ error: 'Failed to generate HTML preview' });
  }
});

// NEW: List template PDF generation endpoint
router.post('/generate-list-pdf', async (req, res) => {
  try {
    const { listHtml, filename, showEmptyCells = true } = req.body;
    
    if (!listHtml || !filename) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    // Create a unique filename
    const pdfFilename = `${filename.replace(/[^a-z0-9]/gi, '_')}_${Date.now()}.pdf`;
    const outputPath = path.join(tempDir, pdfFilename);
    
    // Set up browser for PDF generation
    const browser = await puppeteer.launch({
      headless: 'new',
      args: [
        '--no-sandbox', 
        '--disable-setuid-sandbox',
        '--disable-web-security',
        '--font-render-hinting=none'
      ]
    });
    
    const page = await browser.newPage();
    
    // Generate HTML with styles for list template
    const fullHtml = generateListPdfContent(listHtml, filename, showEmptyCells);
    
    await page.setContent(fullHtml, { waitUntil: 'networkidle0' });
    
    // Emulate screen media instead of print to match preview
    await page.emulateMediaType('screen');
    
    // Configure PDF options
    const pdfOptions = {
      path: outputPath,
      format: 'A4',
      printBackground: true,
      preferCSSPageSize: false,
      margin: {
        top: '20mm',
        right: '20mm', 
        bottom: '20mm',
        left: '20mm'
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
    console.error('List PDF generation error:', error);
    res.status(500).json({ error: 'Failed to generate list PDF' });
  }
});

// NEW: List template HTML preview endpoint
router.post('/preview-list-html', async (req, res) => {
  try {
    const { listHtml, filename, showEmptyCells = true } = req.body;
    
    if (!listHtml || !filename) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    // Generate HTML with styles for list template - exactly the same as PDF
    const fullHtml = generateListPdfContent(listHtml, filename, showEmptyCells);
    
    // Send the HTML
    res.send(fullHtml);
  } catch (error) {
    console.error('List HTML preview generation error:', error);
    res.status(500).json({ error: 'Failed to generate list HTML preview' });
  }
});

module.exports = router; 