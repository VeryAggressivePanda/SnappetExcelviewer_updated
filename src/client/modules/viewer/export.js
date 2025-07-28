/**
 * Export Module - PDF Export & Preview Functionality
 * Verantwoordelijk voor: PDF generation, preview modal, export filtering
 */

// Utility: Remove parent references recursively from node tree
function stripParentReferences(node) {
  if (Array.isArray(node)) {
    node.forEach(stripParentReferences);
    return;
  }
  if (node && typeof node === 'object') {
    delete node.parent;
    if (Array.isArray(node.children)) {
      node.children.forEach(stripParentReferences);
    }
  }
}

// Helper die ALTIJD alle parent-references verwijdert, diep recursief
function deepCloneWithoutParent(obj) {
  if (Array.isArray(obj)) {
    return obj.map(deepCloneWithoutParent);
  } else if (obj && typeof obj === 'object') {
    const clone = {};
    for (const key in obj) {
      if (key !== 'parent') {
        clone[key] = deepCloneWithoutParent(obj[key]);
      }
    }
    return clone;
  }
  return obj;
}

// Helper die circular references detecteert en vervangt
function safeStringify(obj) {
  const seen = new WeakSet();
  return JSON.stringify(obj, function(key, value) {
    if (typeof value === "object" && value !== null) {
      if (seen.has(value)) return "[Circular]";
      seen.add(value);
    }
    return value;
  });
}

// Function to handle PDF export
async function exportToPdf() {
  try {
    const activeSheetId = window.excelData.activeSheetId;
    if (!activeSheetId || !window.excelData.sheetsLoaded[activeSheetId]) {
      alert('Please select a sheet first');
      return;
    }
    
    const { exportColumnSelect, exportPdfButton } = window.ExcelViewerCore.getDOMElements();
    const selectedValue = exportColumnSelect.value;
    if (!selectedValue) {
      alert('Please select an item to export');
      return;
    }
    
    // Get the data for the active sheet
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    
    // Filter the data by the selected top-level value
    let filteredData = { ...sheetData };
    if (selectedValue) {
      // Get nodes that match the selected value
      const matchingNodes = sheetData.root.children.filter(node => node.value === selectedValue);
      
      // Create a deep clone of the matching nodes to avoid modifying the original data
      const clonedNodes = deepCloneWithoutParent(matchingNodes);
      
      // Use the cloned nodes without level normalization to preserve negative levels
      filteredData = {
        ...sheetData,
        root: {
          ...sheetData.root,
          children: clonedNodes
        }
      };
      
      // Strip parent references vóór export/preview
      stripParentReferences(clonedNodes);
      // Debug: check of parent-property is verwijderd
      console.log("Na strippen, parent in eerste node?", 'parent' in clonedNodes[0]);
      
      // Log what we're sending for debugging
      console.log("[PDF EXPORT] Sending to PDF export:", {
        selectedValue,
        nodeCount: filteredData.root.children.length,
        firstNode: filteredData.root.children[0],
        levels: filteredData.root.children.map(n => ({ value: n.value, level: n.level }))
      });
    }
    
    // Create a title for the PDF
    const sheetTitle = document.querySelector(`.tab-button[data-sheet-id="${activeSheetId}"]`).textContent;
    const pdfTitle = `${sheetTitle} - ${selectedValue}`;
    
    // Show loading state
    const originalText = exportPdfButton.textContent;
    exportPdfButton.textContent = 'Generating...';
    exportPdfButton.disabled = true;
    
    // Verwijder parent-references diep uit de hele boom (begin bij root!)
    stripParentReferences(filteredData.root);
    // Gebruik deepCloneWithoutParent om gegarandeerd alle parent-references te strippen
    const cleanData = deepCloneWithoutParent(filteredData.root.children);
    console.log("DEBUG: Data vóór safeStringify", cleanData);
    const cleanDataString = safeStringify(cleanData);
    console.log("[PDF EXPORT] Clean data safe-stringified:", cleanDataString);
    
    // Get empty cells visibility setting
    const { showEmptyCellsCheckbox } = window.ExcelViewerCore.getDOMElements();
    const showEmptyCells = showEmptyCellsCheckbox ? showEmptyCellsCheckbox.checked : true;
    


    let response;
    try {
      response = await fetch('/pdf/generate-pdf', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          html: cleanDataString, // Stuur als string
          filename: pdfTitle,
          showEmptyCells: showEmptyCells
        })
      });
    } catch (fetchErr) {
      console.error('[PDF EXPORT] Fetch error:', fetchErr);
      alert(`Error contacting server: ${fetchErr.message}`);
      exportPdfButton.textContent = originalText;
      exportPdfButton.disabled = false;
      return;
    }
    
    if (!response.ok) {
      let errorText = await response.text();
      console.error(`[PDF EXPORT] PDF generation failed: ${response.statusText}`, errorText);
      alert(`PDF generation failed: ${response.statusText}\n${errorText}`);
      exportPdfButton.textContent = originalText;
      exportPdfButton.disabled = false;
      return;
    }
    
    // Reset button state
    exportPdfButton.textContent = originalText;
    exportPdfButton.disabled = false;
    
    // Get the PDF as a blob
    let blob;
    try {
      blob = await response.blob();
      console.log('[PDF EXPORT] PDF blob received:', blob);
    } catch (blobErr) {
      console.error('[PDF EXPORT] Error reading PDF blob:', blobErr);
      alert(`Error reading PDF blob: ${blobErr.message}`);
      return;
    }
    
    // Create a URL for the blob
    const url = URL.createObjectURL(blob);
    
    // Create a link and click it to download the PDF directly
    const a = document.createElement('a');
    a.href = url;
    a.download = `${pdfTitle}.pdf`;
    document.body.appendChild(a);
    a.click();
    
    // Clean up
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    console.log('[PDF EXPORT] PDF download triggered.');
    
  } catch (error) {
    console.error('[PDF EXPORT] Error generating PDF:', error);
    alert(`Error generating PDF: ${error.message}`);
    const { exportPdfButton } = window.ExcelViewerCore.getDOMElements();
    exportPdfButton.textContent = 'Export to PDF';
    exportPdfButton.disabled = false;
  }
}

// Function to handle PDF preview
function previewPdf() {
  try {
    const activeSheetId = window.excelData.activeSheetId;
    if (!activeSheetId || !window.excelData.sheetsLoaded[activeSheetId]) {
      alert('Please select a sheet first');
      return;
    }
    
    const { exportColumnSelect } = window.ExcelViewerCore.getDOMElements();
    const selectedValue = exportColumnSelect.value;
    if (!selectedValue) {
      alert('Please select an item to preview');
      return;
    }
    
    // Get the data for the active sheet
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    
    // Filter the data by the selected top-level value
    let filteredData = { ...sheetData };
    if (selectedValue) {
      // Get nodes that match the selected value
      const matchingNodes = sheetData.root.children.filter(node => node.value === selectedValue);
      
      // Create a deep clone of the matching nodes to avoid modifying the original data
      const clonedNodes = deepCloneWithoutParent(matchingNodes);
      
      // Use the cloned nodes without level normalization to preserve negative levels
      filteredData = {
        ...sheetData,
        root: {
          ...sheetData.root,
          children: clonedNodes
        }
      };
      
      // Strip parent references vóór export/preview
      stripParentReferences(clonedNodes);
      // Debug: check of parent-property is verwijderd
      console.log("Na strippen, parent in eerste node?", 'parent' in clonedNodes[0]);
      
      // Log what we're sending for debugging
      console.log("Sending to PDF preview:", {
        selectedValue,
        nodeCount: filteredData.root.children.length,
        firstNode: filteredData.root.children[0],
        levels: filteredData.root.children.map(n => ({ value: n.value, level: n.level }))
      });
    }
    
    // Create a title for the preview
    const sheetTitle = document.querySelector(`.tab-button[data-sheet-id="${activeSheetId}"]`).textContent;
    const previewTitle = `${sheetTitle} - ${selectedValue}`;
    
    // Request the exact PDF HTML from server to maintain 100% consistency with PDF
    // Get empty cells visibility setting
    const { showEmptyCellsCheckbox } = window.ExcelViewerCore.getDOMElements();
    const showEmptyCells = showEmptyCellsCheckbox ? showEmptyCellsCheckbox.checked : true;
    


    // Gebruik deepCloneWithoutParent en safeStringify om circular structure errors te voorkomen
    const cleanData = deepCloneWithoutParent(filteredData.root.children);
    const cleanDataString = safeStringify(cleanData);
    console.log("[PDF PREVIEW] Clean data safe-stringified:", cleanDataString);
    
    fetch('/pdf/preview-html', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        html: cleanDataString,
        filename: previewTitle,
        showEmptyCells: showEmptyCells
      })
    })
    .then(response => {
      if (!response.ok) {
        throw new Error('Error getting PDF preview HTML');
      }
      return response.text();
    })
    .then(htmlContent => {
      showExactPdfPreview(htmlContent, previewTitle);
    })
    .catch(error => {
      console.error('Error generating preview:', error);
      alert(`Error generating preview: ${error.message}`);
    });
  } catch (error) {
    console.error('Error generating preview:', error);
    alert(`Error generating preview: ${error.message}`);
  }
}

// Function to show the exact PDF HTML preview in a modal
function showExactPdfPreview(htmlContent, title) {
  // Create modal container
  const modal = document.createElement('div');
  modal.className = 'pdf-preview-modal';
  
  // Create content container
  const content = document.createElement('div');
  content.className = 'pdf-preview-content';
  
  // Add close button
  const closeButton = document.createElement('button');
  closeButton.className = 'pdf-preview-close';
  closeButton.innerHTML = '&times;';
  closeButton.addEventListener('click', () => {
    document.body.removeChild(modal);
  });
  
  // Add title
  const titleElement = document.createElement('h2');
  titleElement.textContent = 'PDF Preview: ' + title;
  titleElement.style.margin = '0 0 15px 0';
  titleElement.style.color = '#34a3d7';
  
  // Add buttons container
  const buttonsContainer = document.createElement('div');
  buttonsContainer.style.display = 'flex';
  buttonsContainer.style.justifyContent = 'space-between';
  buttonsContainer.style.margin = '10px 0';
  buttonsContainer.style.gap = '10px';
  
  // Add left button group
  const leftButtons = document.createElement('div');
  leftButtons.style.display = 'flex';
  leftButtons.style.gap = '10px';
  
  // Add "Inspect Element" hint
  const inspectHint = document.createElement('div');
  inspectHint.textContent = 'Right-click → Inspect to examine the PDF HTML structure';
  inspectHint.style.color = '#666';
  inspectHint.style.fontSize = '0.9rem';
  inspectHint.style.fontStyle = 'italic';
  inspectHint.style.display = 'flex';
  inspectHint.style.alignItems = 'center';
  
  leftButtons.appendChild(inspectHint);
  
  // Add export button
  const exportButton = document.createElement('button');
  exportButton.className = 'control-button';
  exportButton.textContent = 'Export to PDF';
  exportButton.addEventListener('click', () => {
    exportToPdf();
    document.body.removeChild(modal);
  });
  
  buttonsContainer.appendChild(leftButtons);
  buttonsContainer.appendChild(exportButton);
  
  // Create iframe for exact PDF rendering
  const iframe = document.createElement('iframe');
  iframe.style.width = '100%';
  iframe.style.height = 'calc(100% - 100px)';
  iframe.style.border = '1px solid #ddd';
  iframe.style.boxShadow = '0 0 10px rgba(0,0,0,0.1)';
  iframe.style.backgroundColor = 'white';
  
  // Add A4 paper styling to make preview match PDF dimensions
  iframe.style.maxWidth = '210mm';
  iframe.style.margin = '0 auto';
  iframe.style.padding = '1cm';
  iframe.style.display = 'block';
  
  // Assemble the modal
  content.appendChild(closeButton);
  content.appendChild(titleElement);
  content.appendChild(buttonsContainer);
  content.appendChild(iframe);
  modal.appendChild(content);
  
  // Add to document
  document.body.appendChild(modal);
  
  // Write content to iframe after it's in the DOM
  setTimeout(() => {
    const iframeDoc = iframe.contentDocument || iframe.contentWindow.document;
    iframeDoc.open();
    iframeDoc.write(htmlContent);
    iframeDoc.close();
    
    // Fix any potential styling issues in the iframe
    const iframeStyle = document.createElement('style');
    iframeStyle.textContent = `
      @page {
        size: A4;
        margin: 1cm;
      }
      body {
        width: 100%;
        margin: 0;
        padding: 0;
      }
      .pdf-container {
        width: 100%;
        max-width: 100%;
      }
    `;
    iframeDoc.head.appendChild(iframeStyle);
  }, 100);
}

// Export functions for other modules
window.ExcelViewerExport = {
  exportToPdf,
  previewPdf,
  showExactPdfPreview
}; 