document.addEventListener('DOMContentLoaded', function() {
  // DOM Elements
  const tabButtons = document.querySelectorAll('.tab-button');
  const sheetContents = document.querySelectorAll('.sheet-content');
  const fontSizeSelect = document.getElementById('font-size');
  const themeSelect = document.getElementById('theme');
  const cellPaddingSelect = document.getElementById('cell-padding');
  const showGridCheckbox = document.getElementById('show-grid');
  const exportColumnSelect = document.getElementById('export-column');
  const previewPdfButton = document.getElementById('preview-pdf');
  const exportPdfButton = document.getElementById('export-pdf');
  const collapseAllButton = document.getElementById('collapse-all');
  const expandAllButton = document.getElementById('expand-all');
  
  // Global data store - preserve existing properties but ensure all required properties exist
  window.excelData = window.excelData || {};
  window.excelData.fileId = window.excelData.fileId || null;
  window.excelData.sheetsLoaded = window.excelData.sheetsLoaded || {};
  window.excelData.hierarchyConfigurations = window.excelData.hierarchyConfigurations || {};
  window.excelData.activeSheetId = null;
  window.excelData.savedConfigurations = {};
  
  // Load saved hierarchy configurations from localStorage
  try {
    const savedConfigurations = localStorage.getItem('hierarchyConfigurations');
    if (savedConfigurations) {
      const parsed = JSON.parse(savedConfigurations);
      if (parsed && typeof parsed === 'object') {
        console.log('Loaded saved hierarchy configurations from localStorage:', parsed);
        window.excelData.savedConfigurations = parsed;
      }
    } else {
      console.log('No saved hierarchy configurations found in localStorage');
    }
  } catch (error) {
    console.error('Error loading saved configurations:', error);
  }
  
  // Only set fileId if not already set via the script tag in the HTML
  if (!window.excelData.fileId) {
    const fileId = document.body.getAttribute('data-file-id');
    if (fileId) {
      window.excelData.fileId = fileId;
      console.log(`Set file ID from data attribute: ${fileId}`);
    } else {
      console.error('No file ID found in page data attributes');
    }
  } else {
    console.log(`Using existing file ID: ${window.excelData.fileId}`);
  }
  
  // Function to populate export column dropdown
  function populateExportColumnDropdown() {
    const activeSheetId = window.excelData.activeSheetId;
    const exportColumnSelect = document.getElementById('export-column');
    
    // Clear existing options
    exportColumnSelect.innerHTML = '';
    
    // Disable and clear if no active sheet or sheets not loaded
    if (!activeSheetId || !window.excelData.sheetsLoaded[activeSheetId]) {
      exportColumnSelect.disabled = true;
      exportPdfButton.disabled = true;
      previewPdfButton.disabled = true;
      return;
    }
    
    // Enable dropdown if active sheet is valid
    exportColumnSelect.disabled = false;
    exportPdfButton.disabled = false;
    previewPdfButton.disabled = false;
    
    // Get node values from active sheet data
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    if (!sheetData || !sheetData.root || !sheetData.root.children) {
      return;
    }
    
    // Get top-level nodes (level 0 nodes which are the root nodes in each tab)
    const topLevelNodes = sheetData.root.children.filter(node => 
      node.columnIndex === 0
    );
    
    console.log("Top level nodes for export:", topLevelNodes.map(n => ({ value: n.value, level: n.level })));
    
    // Add default option
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.text = 'Select an item to export...';
    exportColumnSelect.appendChild(defaultOption);
    
    // Add option for each top-level node
    topLevelNodes.forEach(node => {
      const option = document.createElement('option');
      option.value = node.value;
      option.text = node.value;
      exportColumnSelect.appendChild(option);
    });
    
    // Try to restore saved selection from localStorage
    const savedSelection = localStorage.getItem(`exportSelection_${activeSheetId}`);
    if (savedSelection) {
      // Find if the saved value exists in current options
      const exists = Array.from(exportColumnSelect.options).some(opt => opt.value === savedSelection);
      if (exists) {
        exportColumnSelect.value = savedSelection;
        // Also enable the PDF buttons if we've restored a selection
        exportPdfButton.disabled = false;
        previewPdfButton.disabled = false;
      }
    }
  }
  
  // Function to save export selection to localStorage
  function saveExportSelection(sheetId, value) {
    if (sheetId && value) {
      localStorage.setItem(`exportSelection_${sheetId}`, value);
    }
  }
  
  // Function to handle PDF export
  async function exportToPdf() {
    try {
      const activeSheetId = window.excelData.activeSheetId;
      if (!activeSheetId || !window.excelData.sheetsLoaded[activeSheetId]) {
        alert('Please select a sheet first');
        return;
      }
      
      const exportColumnSelect = document.getElementById('export-column');
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
        const clonedNodes = JSON.parse(JSON.stringify(matchingNodes));
        
        // Use the cloned nodes without level normalization to preserve negative levels
        filteredData = {
          ...sheetData,
          root: {
            ...sheetData.root,
            children: clonedNodes
          }
        };
        
        // Log what we're sending for debugging
        console.log("Sending to PDF export:", {
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
      
      // Send data to server for PDF generation
      const response = await fetch('/pdf/generate-pdf', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          html: JSON.stringify(filteredData.root.children),
          filename: pdfTitle
        })
      });
      
      if (!response.ok) {
        throw new Error(`PDF generation failed: ${response.statusText}`);
      }
      
      // Reset button state
      exportPdfButton.textContent = originalText;
      exportPdfButton.disabled = false;
      
      // Get the PDF as a blob
      const blob = await response.blob();
      
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
      
    } catch (error) {
      console.error('Error generating PDF:', error);
      alert(`Error generating PDF: ${error.message}`);
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
      
      const exportColumnSelect = document.getElementById('export-column');
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
        const clonedNodes = JSON.parse(JSON.stringify(matchingNodes));
        
        // Use the cloned nodes without level normalization to preserve negative levels
        filteredData = {
          ...sheetData,
          root: {
            ...sheetData.root,
            children: clonedNodes
          }
        };
        
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
      fetch('/pdf/preview-html', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          html: JSON.stringify(filteredData.root.children),
          filename: previewTitle
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
  
  // Function to show column hierarchy configuration modal
  function showHierarchyConfigModal(headers, callback) {
    // Create modal container
    const modal = document.createElement('div');
    modal.className = 'pdf-preview-modal';
    
    // Create content container
    const content = document.createElement('div');
    content.className = 'pdf-preview-content';
    content.style.maxWidth = '600px';
    content.style.padding = '25px';
    
    // Add header
    const header = document.createElement('h2');
    header.textContent = 'Configure Column Hierarchy';
    header.style.marginBottom = '20px';
    header.style.color = '#34a3d7';
    
    // Add description
    const description = document.createElement('p');
    description.textContent = 'Select parent-child relationships between columns to create a hierarchical view.';
    description.style.marginBottom = '20px';
    
    // Check for saved configuration to apply to the form
    const activeSheetId = window.excelData.activeSheetId;
    const configKey = `${window.excelData.fileId}-${activeSheetId}`;
    const savedConfig = window.excelData.savedConfigurations[configKey];
    
    // Create form
    const form = document.createElement('div');
    form.style.display = 'flex';
    form.style.flexDirection = 'column';
    form.style.gap = '15px';
    form.style.marginBottom = '25px';
    
    // For each column (except first), create a dropdown to select its parent
    const selects = {};
    headers.forEach((header, index) => {
      if (!header) return; // Skip empty headers
      
      const group = document.createElement('div');
      group.style.display = 'flex';
      group.style.alignItems = 'center';
      group.style.gap = '10px';
      
      const label = document.createElement('label');
      label.textContent = `Column ${index + 1}: ${header}`;
      label.style.flex = '1';
      label.style.fontWeight = 'bold';
      
      const select = document.createElement('select');
      select.id = `parent-${index}`;
      select.style.padding = '8px';
      select.style.borderRadius = '4px';
      select.style.minWidth = '200px';
      
      // Add "None" option (top-level)
      const noneOption = document.createElement('option');
      noneOption.value = '';
      noneOption.textContent = 'None (Top level)';
      select.appendChild(noneOption);
      
      // Add "Ignore" option
      const ignoreOption = document.createElement('option');
      ignoreOption.value = 'ignore';
      ignoreOption.textContent = 'Ignore this column';
      select.appendChild(ignoreOption);
      
      // Add all previous columns as potential parents
      headers.slice(0, index).forEach((parentHeader, parentIndex) => {
        if (!parentHeader) return; // Skip empty headers
        
        const option = document.createElement('option');
        option.value = parentIndex.toString();
        option.textContent = parentHeader;
        
        select.appendChild(option);
      });
      
      // If saved config exists, select the appropriate value, otherwise default to previous column
      if (savedConfig && savedConfig[index] !== undefined) {
        if (savedConfig[index] === null) {
          select.value = ''; // None (top level)
        } else if (savedConfig[index] === 'ignore') {
          select.value = 'ignore';
        } else {
          select.value = savedConfig[index].toString();
        }
      } else if (index > 0) {
        // Default to previous column if no saved config
        select.value = (index - 1).toString();
      }
      
      group.appendChild(label);
      group.appendChild(select);
      form.appendChild(group);
      
      selects[index] = select;
    });
    
    // Add buttons
    const buttonContainer = document.createElement('div');
    buttonContainer.style.display = 'flex';
    buttonContainer.style.justifyContent = 'space-between';
    buttonContainer.style.gap = '10px';
    
    const leftButtonGroup = document.createElement('div');
    leftButtonGroup.style.display = 'flex';
    leftButtonGroup.style.gap = '10px';
    
    const defaultButton = document.createElement('button');
    defaultButton.textContent = 'Use Default Hierarchy';
    defaultButton.className = 'control-button';
    defaultButton.style.backgroundColor = '#f0f0f0';
    defaultButton.style.color = '#333';
    
    // Add option to delete saved configuration
    const clearSavedButton = document.createElement('button');
    clearSavedButton.textContent = 'Clear Saved Configuration';
    clearSavedButton.className = 'control-button';
    clearSavedButton.style.backgroundColor = '#f8d7da';
    clearSavedButton.style.color = '#721c24';
    
    // Only show the clear button if there's a saved configuration for this file/sheet
    if (!window.excelData.savedConfigurations[configKey]) {
      clearSavedButton.style.display = 'none';
    }
    
    const saveButton = document.createElement('button');
    saveButton.textContent = 'Apply & Save Configuration';
    saveButton.className = 'control-button';
    
    leftButtonGroup.appendChild(defaultButton);
    leftButtonGroup.appendChild(clearSavedButton);
    
    buttonContainer.appendChild(leftButtonGroup);
    buttonContainer.appendChild(saveButton);
    
    // Add all elements to modal
    content.appendChild(header);
    content.appendChild(description);
    content.appendChild(form);
    content.appendChild(buttonContainer);
    modal.appendChild(content);
    
    // Add to document
    document.body.appendChild(modal);
    
    // Handle delete saved configuration button click
    clearSavedButton.addEventListener('click', () => {
      const configKey = `${window.excelData.fileId}-${activeSheetId}`;
      delete window.excelData.savedConfigurations[configKey];
      localStorage.setItem('hierarchyConfigurations', JSON.stringify(window.excelData.savedConfigurations));
      clearSavedButton.textContent = 'Configuration Cleared';
      clearSavedButton.disabled = true;
      setTimeout(() => {
        clearSavedButton.textContent = 'Clear Saved Configuration';
        clearSavedButton.disabled = false;
      }, 1500);
    });
    
    // Handle default button click
    defaultButton.addEventListener('click', () => {
      // Create default hierarchy (linear parent-child relationship)
      const hierarchy = {};
      headers.forEach((header, index) => {
        if (index === 0) {
          hierarchy[index] = null; // Top level
        } else {
          hierarchy[index] = index - 1; // Previous column is parent
        }
      });
      
      document.body.removeChild(modal);
      callback(hierarchy);
    });
    
    // Handle save button click
    saveButton.addEventListener('click', () => {
      const hierarchy = {};
      headers.forEach((header, index) => {
        if (!header) return;
        
        const select = selects[index];
        if (!select) return;
        
        const value = select.value;
        if (value === 'ignore') {
          hierarchy[index] = 'ignore';
        } else if (value === '') {
          hierarchy[index] = null; // Top level
        } else {
          hierarchy[index] = parseInt(value); // Parent column index
        }
      });
      
      document.body.removeChild(modal);
      callback(hierarchy);
    });
  }
  
  // Process raw Excel data into hierarchical structure
  function processExcelData(rawData, hierarchy) {
    console.log("Using hierarchy configuration:", hierarchy);
    
    const headers = rawData.headers;
    const data = rawData.data;
    const root = { children: [] };
    
    // Find top-level columns (those with null parent)
    const topLevelColumns = [];
    for (const colIndex in hierarchy) {
      if (hierarchy[colIndex] === null) {
        topLevelColumns.push(parseInt(colIndex));
      }
    }
    
    // Create a parent-child mapping
    const childrenMap = {};
    for (const colIndex in hierarchy) {
      const parentIndex = hierarchy[colIndex];
      if (parentIndex !== null && parentIndex !== 'ignore') {
        if (!childrenMap[parentIndex]) {
          childrenMap[parentIndex] = [];
        }
        childrenMap[parentIndex].push(parseInt(colIndex));
      }
    }
    
    // Maps to track unique nodes by value
    const nodesByColumn = {};
    
    // Track current parent values for each column level
    const currentValues = {};
    
    // Process each row
    for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
      const row = data[rowIndex];
      if (!row) continue;
      
      // Store current values for each column in this row
      for (const colIndex in hierarchy) {
        const colIdx = parseInt(colIndex);
        if (row[colIdx] !== undefined && row[colIdx] !== null && row[colIdx] !== '') {
          currentValues[colIdx] = row[colIdx];
        }
      }
      
      // Process top-level columns for this row
      for (const topLevelColIndex of topLevelColumns) {
        const value = currentValues[topLevelColIndex] || '';
        if (!value) continue; // Skip if no value
        
        const key = `col${topLevelColIndex}-${value}`;
        
        // Create node if not exists
        if (!nodesByColumn[key]) {
          // Create Excel coordinates reference (A1, B2, etc.)
          const excelColumn = String.fromCharCode(65 + topLevelColIndex);
          const excelRow = rowIndex + 1; // Excel is 1-indexed
          const excelCoord = `${excelColumn}${excelRow}`;
          
          nodesByColumn[key] = {
            id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
            value: value,
            columnName: headers[topLevelColIndex] || `Column ${topLevelColIndex+1}`,
            columnIndex: topLevelColIndex,
            level: 0,
            children: [],
            childNodes: {},
            properties: [],
            // Store Excel coordinates with the node
            excelCoordinates: {
              column: excelColumn,
              row: excelRow,
              cell: excelCoord,
              rowIndex: rowIndex
            }
          };
          
          root.children.push(nodesByColumn[key]);
        }
        
        // Process children recursively with the stored values for this row
        processChildrenWithCurrentValues(
          nodesByColumn[key], 
          row, 
          0, 
          topLevelColIndex, 
          rowIndex, 
          childrenMap, 
          headers, 
          hierarchy,
          currentValues
        );
      }
    }
    
    // Cleanup temporary objects
    root.children.forEach(cleanupNode);
    
    console.log("Processed data structure:", root);
    
    return {
      headers: headers,
      root: root
    };
  }
  
  // Helper function to process children recursively with current values
  function processChildrenWithCurrentValues(
    parentNode, 
    row, 
    level, 
    parentColIndex, 
    rowIndex, 
    childrenMap, 
    headers, 
    hierarchy, 
    currentValues
  ) {
    // Get child columns for this parent
    const childColumns = childrenMap[parentColIndex] || [];
    
    // Process each child column - even if empty
    for (const childColIndex of childColumns) {
      // Get value from currentValues, or use empty string (don't skip empty cells)
      const childValue = currentValues[childColIndex] || '';
      
      // DO NOT SKIP EMPTY CELLS - Include them as empty children with proper labeling
      
      // Special handling for Les and its children
      let childKey;
      
      // Special case: If parent is Les (column 3/D) and child is in columns 4+ (E+),
      // the child should be a DIRECT child of Les, not a sibling.
      // Les nodes are typically at level 2 (or 3 if Course is present)
      const isLesColumnIndex = 
        (headers[parentColIndex] === 'Les') || 
        (parentColIndex === 2) || 
        (parentColIndex === 3); // Support both 0-indexed and 1-indexed Les columns
      
      const isChildOfLes = isLesColumnIndex && childColIndex > parentColIndex;
      
      // For hierarchy columns (main structure), use consistent keys without row index
      if (headers[childColIndex] === 'Blok' || 
          headers[childColIndex] === 'Week' || 
          headers[childColIndex] === 'Les') {
        childKey = `${parentNode.id}-col${childColIndex}-${childValue || 'empty'}`;
        console.log(`Creating HIERARCHY node key: ${childKey}`);
      }
      // For Les children (content columns E, F, G, etc), include row for uniqueness
      else if (isChildOfLes) {
        childKey = `${parentNode.id}-col${childColIndex}-row${rowIndex}-${childValue || 'empty'}`;
        console.log(`Creating LES CHILD node key: ${childKey}`);
      }
      // For other content columns
      else {
        childKey = `${parentNode.id}-col${childColIndex}-row${rowIndex}-${childValue || 'empty'}`;
        console.log(`Creating CONTENT node key: ${childKey}`);
      }
      
      let childNode = parentNode.childNodes[childKey];
      
      if (!childNode) {
        // Create Excel coordinates reference (A1, B2, etc.)
        const excelColumn = String.fromCharCode(65 + childColIndex);
        const excelRow = rowIndex + 1; // Excel is 1-indexed
        const excelCoord = `${excelColumn}${excelRow}`;
        
        // Generate ID differently for hierarchy vs content nodes
        let nodeId;
        // For main hierarchy columns, use consistent IDs
        if (headers[childColIndex] === 'Blok' || 
            headers[childColIndex] === 'Week' || 
            headers[childColIndex] === 'Les') {
          nodeId = `node-${childColIndex}-${childValue.replace(/[^a-zA-Z0-9]/g, '-')}`;
        } 
        // For content nodes or Les children, include row to keep unique
        else {
          nodeId = `node-${childColIndex}-${childValue.replace(/[^a-zA-Z0-9]/g, '-')}-row${rowIndex}`;
        }
        
        // Determine the correct level based on relationship
        let nodeLevel;
        if (isChildOfLes) {
          // Direct children of Les should be level+1 (typically level 3 or 4)
          nodeLevel = parentNode.level + 1;
        } else {
          // Normal parent-child relationship
          nodeLevel = level + 1;
        }
        
        childNode = {
          id: nodeId,
          value: childValue || '', // Just use empty string instead of "Empty [column name]"
          columnName: headers[childColIndex] || `Column ${childColIndex+1}`,
          columnIndex: childColIndex,
          level: nodeLevel,
          children: [],
          properties: [],
          childNodes: {}, // For its own children
          isEmpty: !childValue, // Track if this is an empty cell
          // Store Excel coordinates with the node
          excelCoordinates: {
            column: excelColumn,
            row: excelRow,
            cell: excelCoord,
            rowIndex: rowIndex
          }
        };
        parentNode.children.push(childNode);
        parentNode.childNodes[childKey] = childNode;
      }
      
      // Process its children recursively
      processChildrenWithCurrentValues(
        childNode, 
        row, 
        level + 1, 
        childColIndex, 
        rowIndex, 
        childrenMap, 
        headers, 
        hierarchy,
        currentValues
      );
      
      // Add properties to the node from this specific row
      // We need to do this for each row, even for existing nodes
      if (!childrenMap[childColIndex] || childNode.children.length === 0) {
        addPropertiesToNode(childNode, row, headers, hierarchy);
      }
    }
    
    // Only add properties directly to the parent node for leaf nodes
    if (!childrenMap[parentColIndex] || parentNode.children.length === 0) {
      addPropertiesToNode(parentNode, row, headers, hierarchy);
    }
  }
  
  // Helper to remove temporary objects
  function cleanupNode(node) {
    delete node.childNodes;
    
    for (const child of node.children) {
      cleanupNode(child);
    }
  }
  
  // Helper to add properties
  function addPropertiesToNode(node, row, headers, hierarchy) {
    // Add ALL columns as properties that aren't part of the hierarchy
    for (let i = 0; i < headers.length; i++) {
      // Skip hierarchy columns, but include empty headers (as "Column X")
      if (hierarchy[i] !== undefined) continue;
      
      const columnName = headers[i] || `Column ${i+1}`;
      
      // Skip "Column 9" and "Column 10" as requested
      if (columnName === "Column 9" || columnName === "Column 10") continue;
      
      // Get value (could be empty string)
      const value = (i < row.length) ? (row[i] || '') : '';
      
      // Create Excel coordinates for this property
      const excelColumn = String.fromCharCode(65 + i);
      const excelRow = node.excelCoordinates ? node.excelCoordinates.row : null;
      const excelCoord = excelRow ? `${excelColumn}${excelRow}` : null;
      
      // Add property if not already exists with this column index
      const existingProp = node.properties.find(p => p.columnIndex === i);
      if (!existingProp) {
        node.properties.push({
          columnIndex: i,
          columnName: columnName,
          value: value,
          // Store Excel coordinates with property
          excelCoordinates: excelCoord ? {
            column: excelColumn,
            row: excelRow,
            cell: excelCoord
          } : null
        });
      }
    }
  }
  
  // Tab switching with lazy loading
  tabButtons.forEach(button => {
    button.addEventListener('click', async () => {
      const sheetId = button.getAttribute('data-sheet-id');
      window.excelData.activeSheetId = sheetId;
      
      // Update active tab
      tabButtons.forEach(btn => btn.classList.remove('active'));
      button.classList.add('active');
      
      // Show the selected sheet content
      sheetContents.forEach(content => {
        if (content.id === `sheet-${sheetId}`) {
          content.classList.add('active');
          
          // Check if we need to load this sheet
          if (window.excelData && !window.excelData.sheetsLoaded[sheetId]) {
            loadSheetData(sheetId);
          } else {
            // Update export column dropdown with loaded data
            populateExportColumnDropdown();
          }
        } else {
          content.classList.remove('active');
        }
      });
    });
  });
  
  // Function to load sheet data
  async function loadSheetData(sheetId) {
    const fileId = window.excelData.fileId;
    console.log("Loading sheet data for sheet " + sheetId + ", using fileId: " + fileId);
    
    // If already loaded, return the cached data
    if (window.excelData.sheetsLoaded[sheetId]) {
      renderSheet(sheetId, window.excelData.sheetsLoaded[sheetId]);
      populateExportColumnDropdown(); // Make sure dropdown is populated for already loaded sheets
      return;
    }
    
    // Request sheet data from server
    console.log("Requesting: /api/sheet/" + fileId + "/" + sheetId);
    try {
      const response = await fetch("/api/sheet/" + fileId + "/" + sheetId);
      console.log("Response received, parsing JSON...");
      const data = await response.json();
      console.log("API response parsed:", data);
      
      // Store raw data for later use
      if (!window.rawExcelData) window.rawExcelData = {};
      if (data.data?.data) {
        window.rawExcelData[sheetId] = data.data.data;
        console.log("Stored raw Excel data for sheet", sheetId, "length:", data.data.data.length);
      }
      
      // Cache the data
      window.excelData.sheetsLoaded[sheetId] = data.data;
      
      // Check if this is raw data that needs manual structure building
      if (data.data.isRawData) {
        console.log("Received raw Excel data, enabling live editing mode...");
        
        // Store the raw data for later use
        if (!window.rawExcelData) window.rawExcelData = {};
        window.rawExcelData[sheetId] = data.data.data;
        
        // Create an empty structure for manual building
        const emptyStructure = {
          headers: data.data.headers,
          root: {
            type: 'root',
            children: [],
            level: -1
          },
          isEditable: true,
          rawData: data.data.data
        };
        
        // Store the empty structure
        window.excelData.sheetsLoaded[sheetId] = emptyStructure;
        
        // Render the empty sheet with live editing enabled
        renderSheet(sheetId, emptyStructure);
      } else {
        // Legacy support for old processed data
        renderSheet(sheetId, data.data);
      }
      
      // Update the active sheet ID
      window.excelData.activeSheetId = sheetId;
      
      // Always populate the export column dropdown after loading data
      populateExportColumnDropdown();
      
      return data.data;
    } catch (error) {
      console.error("Error loading sheet data:", error);
      return null;
    }
  }
  
  // Function to save hierarchy configuration to localStorage
  function saveHierarchyConfiguration(key, hierarchy) {
    try {
      // Save to in-memory object
      window.excelData.savedConfigurations[key] = hierarchy;
      
      // Save all configurations to localStorage
      localStorage.setItem('hierarchyConfigurations', JSON.stringify(window.excelData.savedConfigurations));
      console.log('Saved hierarchy configuration to localStorage:', key);
    } catch (error) {
      console.error('Error saving configuration to localStorage:', error);
    }
  }
  
  // Function to save current hierarchy configuration
  function saveCurrentHierarchyConfiguration() {
    try {
      const activeSheetId = window.excelData.activeSheetId;
      if (!activeSheetId) {
        console.error('No active sheet selected');
        alert('Please select a sheet before saving hierarchy configuration');
        return;
      }
      
      const configKey = `${window.excelData.fileId}-${activeSheetId}`;
      const currentConfig = window.excelData.hierarchyConfigurations[activeSheetId];
      
      if (!currentConfig) {
        console.error('No hierarchy configuration found for active sheet');
        alert('No hierarchy configuration found to save');
        return;
      }
      
      // Save using existing function
      saveHierarchyConfiguration(configKey, currentConfig);
      
      // Provide user feedback
      alert('Hierarchy configuration saved successfully');
    } catch (error) {
      console.error('Error saving current hierarchy configuration:', error);
      alert(`Error saving configuration: ${error.message}`);
    }
  }
  
  // Render hierarchical data to the sheet
  function renderSheet(sheetId, sheetData) {
    if (!sheetData || !sheetData.root) {
      console.error("No data or root provided to renderSheet");
      return;
    }
    
    // Store current expand/collapse state before rendering
    let expandStateMap = {};
    
    // Check if we have a stored expandStateMap from a recent add/delete operation
    if (window.expandStateMap) {
      expandStateMap = window.expandStateMap;
      window.expandStateMap = null; // Clear after use
      console.log("Using stored expand state map");
    } else {
      // Otherwise get the current state from the DOM
      const sheetContent = document.getElementById(`sheet-${sheetId}`);
      if (sheetContent) {
        const nodes = sheetContent.querySelectorAll('[class*="-node"]');
        nodes.forEach(node => {
          // Get the node ID from its data attribute
          const nodeId = node.getAttribute('data-node-id');
          if (nodeId) {
            // Store if the node is collapsed or expanded
            expandStateMap[nodeId] = !node.classList.contains(`level-${node.getAttribute('data-level')}-collapsed`);
          }
        });
      }
    }
    
    const content = document.getElementById(`sheet-${sheetId}`);
    content.innerHTML = '';  // Clear content
    
    // Create hierarchy container
    const hierarchyContainer = document.createElement('div');
    hierarchyContainer.className = 'hierarchy-container';
    
    // Create controls bar
    const controlsBar = document.createElement('div');
    controlsBar.className = 'controls-bar';
    controlsBar.style.display = 'flex';
    controlsBar.style.alignItems = 'center';
    controlsBar.style.marginBottom = '10px';
    controlsBar.style.padding = '5px';
    controlsBar.style.background = '#f8f9fa';
    controlsBar.style.border = '1px solid #ddd';
    controlsBar.style.borderRadius = '4px';
    
    // Create expand/collapse all buttons
    const expandAllBtn = document.createElement('button');
    expandAllBtn.textContent = 'Expand All';
    expandAllBtn.className = 'expand-all-button';
    expandAllBtn.style.marginRight = '10px';
    expandAllBtn.style.padding = '5px 10px';
    expandAllBtn.style.border = '1px solid #ddd';
    expandAllBtn.style.borderRadius = '4px';
    expandAllBtn.style.cursor = 'pointer';
    
    const collapseAllBtn = document.createElement('button');
    collapseAllBtn.textContent = 'Collapse All';
    collapseAllBtn.className = 'collapse-all-button';
    collapseAllBtn.style.marginRight = '10px';
    collapseAllBtn.style.padding = '5px 10px';
    collapseAllBtn.style.border = '1px solid #ddd';
    collapseAllBtn.style.borderRadius = '4px';
    collapseAllBtn.style.cursor = 'pointer';
    
    // Add click handlers
    expandAllBtn.addEventListener('click', () => toggleAllNodes(false));
    collapseAllBtn.addEventListener('click', () => toggleAllNodes(true));
    
    // Add save hierarchy button
    const saveHierarchyButton = document.createElement('button');
    saveHierarchyButton.textContent = 'Save Hierarchy';
    saveHierarchyButton.className = 'save-hierarchy-button';
    saveHierarchyButton.style.marginLeft = 'auto';
    saveHierarchyButton.style.padding = '5px 10px';
    saveHierarchyButton.style.border = '1px solid #ddd';
    saveHierarchyButton.style.borderRadius = '4px';
    saveHierarchyButton.style.cursor = 'pointer';
    
    saveHierarchyButton.addEventListener('click', () => {
      // Show saving indicator
      const oldText = saveHierarchyButton.textContent;
      saveHierarchyButton.textContent = 'Saving...';
      saveHierarchyButton.disabled = true;
      
      // Simulate async operation
      setTimeout(() => {
        try {
          // Serialize hierarchy
          const hierarchy = {};
          
          saveHierarchyButton.textContent = 'Saved ✓';
          setTimeout(() => {
            saveHierarchyButton.textContent = oldText;
            saveHierarchyButton.disabled = false;
          }, 1500);
        } catch (e) {
          console.error("Error saving hierarchy:", e);
          saveHierarchyButton.textContent = 'Error Saving';
          saveHierarchyButton.style.color = 'red';
          setTimeout(() => {
            saveHierarchyButton.textContent = oldText;
            saveHierarchyButton.style.color = '';
            saveHierarchyButton.disabled = false;
          }, 1500);
        }
      }, 500);
    });
    
    // Add buttons to controls bar
    controlsBar.appendChild(expandAllBtn);
    controlsBar.appendChild(collapseAllBtn);
    controlsBar.appendChild(saveHierarchyButton);
    
    // Add controls bar to container
    hierarchyContainer.appendChild(controlsBar);
    
    // Check if data is empty or if this is a raw data sheet
    if (!sheetData.root.children || sheetData.root.children.length === 0) {
      // Create a default starter node for building
      const starterNode = {
        id: `node-starter-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        value: 'Start Building',
        columnName: '',
        level: 0,
        children: [],
        properties: [],
        isPlaceholder: true
      };
      
      sheetData.root.children = [starterNode];
      
      // Enable live editing for this sheet
      if (window.enableLiveEditing) {
        window.enableLiveEditing(sheetId);
      }
    }
    
    // Enable live editing mode for all sheets
    sheetData.isEditable = true;
    
    // Enable live editing functionality
    if (window.enableLiveEditing) {
      window.enableLiveEditing(sheetId);
    }
    
    // Render each child of the root
    sheetData.root.children.forEach(child => {
      const nodeElement = renderNode(child, 0, expandStateMap);
      if (nodeElement) {
        hierarchyContainer.appendChild(nodeElement);
      }
    });
    
    // Attach content to page
    content.appendChild(hierarchyContainer);
    
    // Apply layout styling after rendering
    if (window.applyLayoutStyling) {
      setTimeout(() => {
        window.applyLayoutStyling();
      }, 100);
    }
  }
  
  // Helper function to truncate text with middle ellipsis
  function truncateMiddle(text, maxLength = 35) {
    if (!text || text.length <= maxLength) return text;
    
    const start = Math.floor(maxLength / 2);
    const end = Math.ceil(maxLength / 2);
    
    return text.substring(0, start) + '...' + text.substring(text.length - end);
  }
  
  function renderNode(node, level, expandStateMap = {}) {
    if (!node) return null;
    
    // Skip hidden nodes
    if (node.hidden === true) {
      return null;
    }
    
    // Use the node's own level if it has one, otherwise use the passed level
    const nodeLevel = node.level !== undefined ? node.level : level;
    
    const nodeEl = document.createElement('div');
    
    // Use level-specific class names
    // Support negative levels by using _1, _2, etc. instead of -1, -2 for CSS compatibility
    const cssLevel = nodeLevel < 0 ? `_${Math.abs(nodeLevel)}` : nodeLevel;
    nodeEl.className = `level-${cssLevel}-node`;
    
    if (node.isEmpty) {
      nodeEl.classList.add(`level-${cssLevel}-empty`);
    }
    
    // Add data attributes for styling and debugging
    nodeEl.setAttribute('data-level', nodeLevel);
    nodeEl.setAttribute('data-column-name', node.columnName || '');
    nodeEl.setAttribute('data-column-index', node.columnIndex || '');
    nodeEl.setAttribute('data-is-empty', node.isEmpty ? 'true' : 'false');
    nodeEl.setAttribute('data-node-id', node.id || '');
    
    // Set initial collapse state based on level and expandStateMap
    const shouldBeExpanded = expandStateMap[node.id] !== undefined ? 
      expandStateMap[node.id] : (nodeLevel >= 2);
    
    if (!shouldBeExpanded) {
      nodeEl.classList.add(`level-${cssLevel}-collapsed`);
    }
    
    // Node header - contains column title AND cell data if this is a parent node
    const header = document.createElement('div');
    header.className = `level-${cssLevel}-header`;
    if (node.isEmpty) {
      header.classList.add(`level-${cssLevel}-empty-header`);
    }
    
    // Add expand button for level _1, 0 and 1
    if (nodeLevel < 2 || nodeLevel === -1) {
      const expandBtn = document.createElement('div');
      expandBtn.className = `level-${cssLevel}-expand-button`;
      header.appendChild(expandBtn);
    }
    
    const title = document.createElement('div');
    title.className = `level-${cssLevel}-title`;
    
    // Use normal ellipsis style for long text
    title.style.overflow = 'hidden';
    title.style.textOverflow = 'ellipsis';
    title.style.whiteSpace = 'nowrap';
    title.style.maxWidth = '100%';
    
    // IMPORTANT: For parent nodes, show both column name AND value in the header
    // Check if node has children to determine if it's a parent
    const isParent = node.children && node.children.length > 0;
    
    if (isParent) {
      // For parent nodes: title shows both column name and cell value with middle ellipsis
      const columnName = node.columnName || '';
      const nodeValue = node.value || '';
      
      // Apply middle ellipsis to the value, keep column name intact
      const truncatedValue = truncateMiddle(nodeValue);
      title.innerHTML = `<span class="column-name">${columnName}:</span> <span class="cell-value">${truncatedValue}</span>`;
      
      // Add tooltip with full text if truncated
      if (nodeValue !== truncatedValue) {
        title.title = `${columnName}: ${nodeValue}`;
      }
    } else {
      // For leaf nodes: title shows just column name
      title.textContent = node.columnName;
    }
    
    header.appendChild(title);
    
    // Use the unified edit controls from live-editor.js
    if (window.addEditControls) {
      window.addEditControls(node, nodeEl, header);
    } else {
      // Fallback if live-editor.js is not loaded
      console.warn('addEditControls not available, using fallback');
    }
    
    // Manage Items button is now handled by addEditControls function
    
    // Add click handler for expanding/collapsing
    header.addEventListener('click', (e) => {
      if (nodeLevel < 3 || nodeLevel === -1) { // Allow clicking for hierarchy elements
        nodeEl.classList.toggle(`level-${cssLevel}-collapsed`);
        e.stopPropagation();
      }
    });
    
    nodeEl.appendChild(header);
    
    // Node content - contains EITHER child nodes OR a cell value (for leaf nodes)
    const content = document.createElement('div');
    content.className = `level-${cssLevel}-content`;
    
    // If node has children, render them
    if (node.children && node.children.length > 0) {
      // For parent nodes: content contains children
      const childrenContainer = document.createElement('div');
      childrenContainer.className = `level-${cssLevel}-children`;
      childrenContainer.setAttribute('data-child-count', node.children.length);
      
      // Add class based on number of children to allow CSS targeting
      childrenContainer.classList.add(`level-${cssLevel}-child-count-${node.children.length}`);
      
      // Add layout class specific to level and number of children
      if (nodeLevel === -1) {
        childrenContainer.classList.add('level-_1-vertical-layout');
      } else if (nodeLevel === 0) {
        childrenContainer.classList.add('level-0-vertical-layout');
      } else if (nodeLevel === 1) {
        // Level 1 uses flexible layout based on child count
        if (node.children.length === 1) {
          childrenContainer.classList.add('level-1-single-child');
        } else if (node.children.length === 2) {
          childrenContainer.classList.add('level-1-two-children');
        } else if (node.children.length === 3) {
          childrenContainer.classList.add('level-1-three-children');
        } else {
          childrenContainer.classList.add('level-1-many-children');
        }
        childrenContainer.classList.add('level-1-flex-layout');
      } else if (nodeLevel === 2) {
        childrenContainer.classList.add('level-2-vertical-layout');
      } else if (nodeLevel === 3) {
        childrenContainer.classList.add('level-3-vertical-layout');
      } else if (nodeLevel === 4) {
        childrenContainer.classList.add('level-4-vertical-layout');
      }
      
      // Render all children with proper level nesting
      node.children.forEach(child => {
        const childEl = renderNode(child, nodeLevel + 1, expandStateMap);
        if (childEl) {
          childrenContainer.appendChild(childEl);
        }
      });
      
      content.appendChild(childrenContainer);
    } else {
      // For leaf nodes: content contains cell value
      content.textContent = node.value || '';
      
      if (node.isEmpty) {
        content.classList.add(`level-${cssLevel}-empty-content`);
      }
    }
    
    nodeEl.appendChild(content);
    
    // Process properties (non-hierarchy columns)
    if (node.properties && node.properties.length > 0) {
      const propsContainer = document.createElement('div');
      propsContainer.className = `level-${cssLevel}-properties`;
      
      // Create a row of property nodes
      node.properties.forEach(prop => {
        // Skip "Column 9" and "Column 10" properties
        if (prop.columnName === "Column 9" || prop.columnName === "Column 10") return;
        
        // Create a property node
        const propLevel = Math.min(nodeLevel + 1, 4); // Cap at level 4 for properties
        const propNode = document.createElement('div');
        propNode.className = `level-${propLevel}-node`;
        propNode.setAttribute('data-column-name', prop.columnName);
        propNode.setAttribute('data-column-index', prop.columnIndex);
        propNode.setAttribute('data-level', propLevel);
        
        // Create property header - contains ONLY column name for properties
        const propHeader = document.createElement('div');
        propHeader.className = `level-${propLevel}-header`;
        
        // Create title for header
        const propTitle = document.createElement('div');
        propTitle.className = `level-${propLevel}-title`;
        propTitle.textContent = prop.columnName;
        propHeader.appendChild(propTitle);
        
        // Create property content - contains cell value
        const propContent = document.createElement('div');
        propContent.className = `level-${propLevel}-content`;
        propContent.textContent = prop.value || '';
        
        if (!prop.value || prop.value.trim() === '') {
          propNode.classList.add(`level-${propLevel}-empty`);
          propHeader.classList.add(`level-${propLevel}-empty-header`);
          propContent.classList.add(`level-${propLevel}-empty-content`);
        }
        
        // Assemble property node
        propNode.appendChild(propHeader);
        propNode.appendChild(propContent);
        propsContainer.appendChild(propNode);
      });
      
      nodeEl.appendChild(propsContainer);
    }
    
    return nodeEl;
  }
  
  // Show modal to add a child element
  function showAddChildElementModal(parentNode, parentLevel) {
    console.log("Show add child modal for:", {
      parentNodeId: parentNode.id,
      parentNodeValue: parentNode.value, 
      parentNodeColumnName: parentNode.columnName,
      parentLevel: parentLevel,
      excelCoordinates: parentNode.excelCoordinates
    });
    
    // Create modal container
    const modal = document.createElement('div');
    modal.className = 'element-modal';
    modal.style.position = 'fixed';
    modal.style.top = '0';
    modal.style.left = '0';
    modal.style.width = '100%';
    modal.style.height = '100%';
    modal.style.backgroundColor = 'rgba(0, 0, 0, 0.5)';
    modal.style.display = 'flex';
    modal.style.justifyContent = 'center';
    modal.style.alignItems = 'center';
    modal.style.zIndex = '1000';
    
    // Create content container
    const content = document.createElement('div');
    content.className = 'element-modal-content';
    content.style.backgroundColor = 'white';
    content.style.padding = '20px';
    content.style.borderRadius = '8px';
    content.style.maxWidth = '500px';
    content.style.width = '90%';
    
    // Create header
    const header = document.createElement('h2');
    header.textContent = `Add Child Element to: ${parentNode.columnName} (${parentNode.value})`;
    header.style.marginBottom = '15px';
    content.appendChild(header);
    
    // Create form
    const form = document.createElement('form');
    form.style.display = 'flex';
    form.style.flexDirection = 'column';
    form.style.gap = '15px';
    
    // Get Excel data
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];

    // Log sheet data keys to help debug
    console.log("Sheet data keys:", Object.keys(sheetData || {}));
    
    // Get the raw data directly from the API response
    const rawData = sheetData?.data?.data || [];
    console.log("Found raw data:", rawData.length > 0 ? "Yes" : "No", "Length:", rawData.length);
    
    // Log the first few rows if available
    if (rawData.length > 0) {
      console.log("Headers row:", rawData[0]);
      if (rawData.length > 1) console.log("First data row:", rawData[1]);
    }
    
    const sheetHeaders = sheetData?.headers || (rawData.length > 0 ? rawData[0] : []);
    
    // Find the correct indices for key columns
    let verlengdeInstructieIndex = -1;
    let zelfstandigeVerwerkingIndex = -1;
    let werkbladIndex = -1;
    
    sheetHeaders.forEach((header, index) => {
      if (header && typeof header === 'string') {
        if (header.includes('Verlengde Instructie')) {
          verlengdeInstructieIndex = index;
          console.log(`Found 'Verlengde Instructie' at index ${index}`);
        }
        if (header.includes('Zelfstandige verwerking')) {
          zelfstandigeVerwerkingIndex = index;
          console.log(`Found 'Zelfstandige verwerking' at index ${index}`);
        }
        if (header === 'Werkblad') {
          werkbladIndex = index;
          console.log(`Found 'Werkblad' at index ${index}`);
        }
      }
    });
    
    // Add simple column selector
    const columnGroup = document.createElement('div');
    const columnLabel = document.createElement('label');
    columnLabel.textContent = 'Select Column:';
    columnLabel.style.fontWeight = 'bold';
    columnLabel.style.marginBottom = '5px';
    columnLabel.style.display = 'block';
    
    const columnSelect = document.createElement('select');
    columnSelect.style.width = '100%';
    columnSelect.style.padding = '8px';
    columnSelect.style.borderRadius = '4px';
    columnSelect.style.border = '1px solid #ced4da';
    
    // Add options for common columns with correct indices
    const columnOptions = [
      { name: 'Verlengde Instructie', index: verlengdeInstructieIndex !== -1 ? verlengdeInstructieIndex : 6 },
      { name: 'Zelfstandige verwerking', index: zelfstandigeVerwerkingIndex !== -1 ? zelfstandigeVerwerkingIndex : 7 },
      { name: 'Werkblad', index: werkbladIndex !== -1 ? werkbladIndex : 3 },
      { name: 'Other', index: -1 }
    ];
    
    columnOptions.forEach(col => {
      if (col.index >= 0 || col.name === 'Other') {
        const option = document.createElement('option');
        option.value = col.index;
        option.textContent = col.name;
        columnSelect.appendChild(option);
      }
    });
    
    // Create a container for the "other" column selector (initially hidden)
    const otherColumnContainer = document.createElement('div');
    otherColumnContainer.style.marginTop = '10px';
    otherColumnContainer.style.display = 'none';
    
    const otherColumnSelect = document.createElement('select');
    otherColumnSelect.style.width = '100%';
    otherColumnSelect.style.padding = '8px';
    otherColumnSelect.style.borderRadius = '4px';
    otherColumnSelect.style.border = '1px solid #ced4da';
    
    // Add all Excel columns to the "other" selector with proper column letters
    sheetHeaders.forEach((header, index) => {
      // Skip parent column and empty headers
      if (index === parentNode.columnIndex || !header) return;
      
      const columnLetter = String.fromCharCode(65 + index);
      const option = document.createElement('option');
      option.value = index;
      option.textContent = `${columnLetter}: ${header}`;
      otherColumnSelect.appendChild(option);
    });
    
    otherColumnContainer.appendChild(otherColumnSelect);
    
    // Show "other" selector if "Other" is selected
    columnSelect.addEventListener('change', () => {
      if (columnSelect.value === '-1') {
        otherColumnContainer.style.display = 'block';
      } else {
        otherColumnContainer.style.display = 'none';
      }
      
      // Update preview
      updatePreview();
    });
    
    otherColumnSelect.addEventListener('change', updatePreview);
    
    // Add to form
    columnGroup.appendChild(columnLabel);
    columnGroup.appendChild(columnSelect);
    columnGroup.appendChild(otherColumnContainer);
    form.appendChild(columnGroup);
    
    // Preview section
    const previewContainer = document.createElement('div');
    previewContainer.style.marginTop = '15px';
    previewContainer.style.padding = '10px';
    previewContainer.style.backgroundColor = '#f8f9fa';
    previewContainer.style.borderRadius = '4px';
    previewContainer.style.fontSize = '0.9rem';
    
    const previewHeader = document.createElement('div');
    previewHeader.textContent = 'Preview';
    previewHeader.style.fontWeight = 'bold';
    previewHeader.style.marginBottom = '5px';
    
    const previewContent = document.createElement('div');
    previewContent.style.fontFamily = 'monospace';
    
    previewContainer.appendChild(previewHeader);
    previewContainer.appendChild(previewContent);
    form.appendChild(previewContainer);
    
    // Update preview with actual Excel data
    function updatePreview() {
      console.log("Updating preview...");
      const selectedColIndex = columnSelect.value === '-1' ? parseInt(otherColumnSelect.value) : parseInt(columnSelect.value);
      
      console.log("Selected column index:", selectedColIndex);
      
      // DIRECT ACCESS APPROACH - don't use complicated coordinate mapping
      // If we have parentNode.excelCoordinates, we can directly access the row
      if (parentNode.excelCoordinates && parentNode.excelCoordinates.rowIndex !== undefined) {
        const rowIndex = parentNode.excelCoordinates.rowIndex;
        
        console.log(`Using direct row index: ${rowIndex} for parent node:`, parentNode.value);
        
        if (rawData && rawData.length > rowIndex && rowIndex >= 0) {
          const row = rawData[rowIndex];
          const excelRow = rowIndex + 1; // Convert to Excel's 1-based row numbering
          const columnLetter = String.fromCharCode(65 + selectedColIndex);
          const cellValue = row && selectedColIndex < row.length ? row[selectedColIndex] : null;
          
          console.log(`Direct access - Found row: ${rowIndex}, column: ${selectedColIndex}, value:`, cellValue);
          
          const coordinates = {
            rowIndex: rowIndex,
            columnIndex: selectedColIndex,
            excelRow: excelRow,
            excelColumn: columnLetter,
            excelCell: `${columnLetter}${excelRow}`,
            value: cellValue
          };
          
          const columnName = sheetHeaders[selectedColIndex] || `Column ${selectedColIndex+1}`;
          
          previewContent.innerHTML = `
            <p><strong>Selected:</strong> ${columnName} (Column ${coordinates.excelColumn})</p>
            <p><strong>Excel Cell:</strong> ${coordinates.excelCell}</p>
            <p><strong>Value:</strong> "${cellValue || '(empty)'}"</p>
          `;
          
          // Allow adding elements with empty values - only disable for null/undefined
          addButton.disabled = cellValue === null || cellValue === undefined;
          
          return;
        } else {
          console.error(`Invalid row index ${rowIndex} for raw data of length ${rawData?.length}`);
        }
      }
      
      // Fallback to regular getExcelCoordinates
      console.log("Falling back to getExcelCoordinates function");
      const coordinates = getExcelCoordinates(parentNode, selectedColIndex);
      
      if (!coordinates) {
        previewContent.innerHTML = '<p class="error">Could not determine Excel coordinates</p>';
        return;
      }
      
      const cellValue = coordinates.value;
      const columnName = sheetHeaders[selectedColIndex] || `Column ${selectedColIndex+1}`;
      
      previewContent.innerHTML = `
        <p><strong>Selected:</strong> ${columnName} (Column ${coordinates.excelColumn})</p>
        <p><strong>Excel Cell:</strong> ${coordinates.excelCell}</p>
        <p><strong>Value:</strong> "${cellValue || '(empty)'}"</p>
      `;
      
      // Allow adding elements with empty values - only disable for null/undefined
      addButton.disabled = cellValue === null || cellValue === undefined;
    }
    
    // Add buttons
    const buttonContainer = document.createElement('div');
    buttonContainer.style.display = 'flex';
    buttonContainer.style.justifyContent = 'space-between';
    buttonContainer.style.marginTop = '20px';
    
    const cancelButton = document.createElement('button');
    cancelButton.textContent = 'Cancel';
    cancelButton.type = 'button';
    cancelButton.style.padding = '8px 15px';
    
    const addButton = document.createElement('button');
    addButton.textContent = 'Add Element';
    addButton.type = 'submit';
    addButton.style.padding = '8px 15px';
    addButton.style.backgroundColor = '#34a3d7';
    addButton.style.color = 'white';
    addButton.style.border = 'none';
    
    buttonContainer.appendChild(cancelButton);
    buttonContainer.appendChild(addButton);
    form.appendChild(buttonContainer);
    
    // Add form to content
    content.appendChild(form);
    
    // Add content to modal
    modal.appendChild(content);
    
    // Add modal to document
    document.body.appendChild(modal);
    
    // Initial preview update
    updatePreview();
    
    // Handle form submission
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      
      // Get the selected column
      const selectedColIndex = columnSelect.value === '-1' ? parseInt(otherColumnSelect.value) : parseInt(columnSelect.value);
      const columnName = sheetHeaders[selectedColIndex] || `Column ${selectedColIndex+1}`;
      
      // DIRECT ACCESS APPROACH - don't use complicated coordinate mapping
      // If we have parentNode.excelCoordinates, we can directly access the row
      let coordinates = null;
      
      if (parentNode.excelCoordinates && parentNode.excelCoordinates.rowIndex !== undefined) {
        const rowIndex = parentNode.excelCoordinates.rowIndex;
        
        if (rawData && rawData.length > rowIndex && rowIndex >= 0) {
          const row = rawData[rowIndex];
          const excelRow = rowIndex + 1; // Convert to Excel's 1-based row numbering
          const columnLetter = String.fromCharCode(65 + selectedColIndex);
          const cellValue = row && selectedColIndex < row.length ? row[selectedColIndex] : '';
          
          coordinates = {
            rowIndex: rowIndex,
            columnIndex: selectedColIndex,
            excelRow: excelRow,
            excelColumn: columnLetter,
            excelCell: `${columnLetter}${excelRow}`,
            value: cellValue
          };
        }
      }
      
      // Fallback to regular getExcelCoordinates if direct access failed
      if (!coordinates) {
        coordinates = getExcelCoordinates(parentNode, selectedColIndex);
      }
      
      if (!coordinates) {
        alert("Could not determine Excel coordinates for this element");
        return;
      }
      
      console.log(`Using Excel cell ${coordinates.excelCell} for ${columnName}`);
      
      // Get the value from the specified cell
      const value = coordinates.value;
      
      // Allow empty values (empty string) - only block null/undefined
      if (value === null || value === undefined) {
        alert(`No valid data found in Excel cell ${coordinates.excelCell}`);
        return;
      }
      
      // Check if already exists
      if (parentNode.children) {
        const existing = parentNode.children.find(child => 
          child.columnIndex === selectedColIndex && child.value === value
        );
        
        if (existing) {
          alert(`${columnName} already exists as a child element`);
          return;
        }
      }
      
      // Create new node
      const newNode = {
        id: `node-${Date.now()}-${Math.floor(Math.random() * 10000)}`,
        columnName: columnName,
        columnIndex: selectedColIndex,
        value: value,
        level: parentLevel + 1,
        children: [],
        properties: []
      };
      
      // Add properties from the same row
      const targetRow = rawData ? rawData[coordinates.rowIndex] : null;
      
      // Check if targetRow exists before trying to iterate through its properties
      if (targetRow && Array.isArray(targetRow)) {
        for (let i = 0; i < targetRow.length; i++) {
          // Skip hierarchy columns
          if (i === selectedColIndex || i === parentNode.columnIndex) continue;
          
          const propValue = targetRow[i];
          if (propValue && propValue.toString().trim() !== '') {
            newNode.properties.push({
              columnIndex: i,
              columnName: sheetHeaders[i] || `Column ${i+1}`,
              value: propValue
            });
          }
        }
      } else {
        console.warn("Target row not found or not an array, skipping property extraction");
      }
      
      // Close modal
      document.body.removeChild(modal);
      
      // Add node and re-render with preserved expand state
      addNode(parentNode, newNode);
    });
    
    // Handle cancel
    cancelButton.addEventListener('click', () => {
      document.body.removeChild(modal);
    });
  }
  
  // Handle style controls
  function applyStyles() {
    // Font size
    document.body.style.setProperty('--font-size', fontSizeSelect.value);
    
    // Theme
    document.body.className = themeSelect.value + '-theme';
    
    // Cell padding
    document.body.style.setProperty('--cell-padding', cellPaddingSelect.value);
    
    // Grid
    document.body.classList.toggle('no-grid', !showGridCheckbox.checked);
  }
  
  // Function to handle collapse/expand all
  function toggleAllNodes(collapse) {
    const activeSheetId = window.excelData.activeSheetId;
    if (!activeSheetId) return;
    
    const sheetContent = document.getElementById(`sheet-${activeSheetId}`);
    if (!sheetContent) return;
    
    const nodes = sheetContent.querySelectorAll('[class*="-node"]');
    nodes.forEach(node => {
      const header = node.querySelector('[class*="-header"]');
      if (header) {
        if (collapse) {
          node.classList.add('level-0-collapsed');
        } else {
          node.classList.remove('level-0-collapsed');
        }
      }
    });
  }
  
  // Initialize
  function initialize() {
    // Set default active tab
    if (tabButtons.length > 0) {
      tabButtons[0].click();
    }
    
    // Initialize style controls
    fontSizeSelect.addEventListener('change', applyStyles);
    themeSelect.addEventListener('change', applyStyles);
    cellPaddingSelect.addEventListener('change', applyStyles);
    showGridCheckbox.addEventListener('change', applyStyles);
    
    // Initialize export controls
    exportPdfButton.addEventListener('click', exportToPdf);
    previewPdfButton.addEventListener('click', previewPdf);
    
    // Use correct reference for exportColumnSelect
    const exportColumnSelect = document.getElementById('export-column');
    exportColumnSelect.addEventListener('change', function() {
      const selectedValue = this.value;
      exportPdfButton.disabled = !selectedValue;
      previewPdfButton.disabled = !selectedValue;
      
      // Save selection to localStorage
      if (selectedValue) {
        const activeSheetId = window.excelData.activeSheetId;
        saveExportSelection(activeSheetId, selectedValue);
      }
    });
    
    // Initialize toolbar including the Add Parent Level button
    initToolbar();
    
    // Apply initial styles
    applyStyles();
  }
  
  // Simple function to get Excel coordinates
  function getExcelCoordinates(node, targetColumnIndex) {
    console.log("NODE DEBUG:", {
      nodeId: node.id,
      value: node.value,
      columnName: node.columnName,
      level: node.level,
      hasExcelCoords: !!node.excelCoordinates,
      excelCoordinates: node.excelCoordinates,
      targetColumn: targetColumnIndex
    });
    
    // Get the active sheet data
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    
    console.log("Sheet data available:", !!sheetData);
    
    // Try to get raw data from the API response
    let rawData = [];
    if (sheetData?.data?.data && Array.isArray(sheetData.data.data)) {
      rawData = sheetData.data.data;
      console.log("Using data from sheetData.data.data, length:", rawData.length);
    }
    
    // If no data was found, try another approach using the loaded Excel data
    if (!rawData || rawData.length <= 1) {
      const excelFilePath = sheetData?.filePath || window.excelData.filePath;
      console.log("No data found, looking for Excel data from file:", excelFilePath);
      
      // Get the Excel file data from the window object
      const excelFile = window.excelData.loadedFiles?.[excelFilePath];
      if (excelFile && excelFile.sheets && excelFile.sheets[activeSheetId]) {
        rawData = excelFile.sheets[activeSheetId].data || [];
        console.log("Found data in excelFile.sheets, length:", rawData.length);
      } else if (window.rawExcelData && window.rawExcelData[activeSheetId]) {
        rawData = window.rawExcelData[activeSheetId] || [];
        console.log("Found data in window.rawExcelData, length:", rawData.length);
      }
    }
    
    // Final check for raw data
    if (!rawData || rawData.length <= 1) {
      console.error("No Excel data available, rawData length:", rawData?.length);
      
      // EMERGENCY FALLBACK: Get data from any available source
      if (node.excelCoordinates && node.excelCoordinates.rowIndex !== undefined) {
        const lesNumber = parseInt((node.value || '').match(/\d+/)?.[0] || '0');
        console.log("Emergency fallback: Found Les number", lesNumber);
        return {
          rowIndex: lesNumber,
          columnIndex: targetColumnIndex,
          excelRow: lesNumber + 1,
          excelColumn: String.fromCharCode(65 + targetColumnIndex),
          excelCell: `${String.fromCharCode(65 + targetColumnIndex)}${lesNumber + 1}`,
          value: `Unable to load value for ${node.value}`
        };
      }
      return null;
    }
    
    // Log the first few rows of data we found for debugging
    console.log("First row (headers):", JSON.stringify(rawData[0]));
    if (rawData.length > 1) console.log("Second row:", JSON.stringify(rawData[1]));
    if (rawData.length > 2) console.log("Third row:", JSON.stringify(rawData[2]));
    
    // Check if excelCoordinates might be lower in the node hierarchy
    if (!node.excelCoordinates && node.parent && node.parent.excelCoordinates) {
      console.log("Using parent's Excel coordinates as fallback");
      node.excelCoordinates = node.parent.excelCoordinates;
    }
    
    // Direct access approach - use the stored row index directly
    if (node.excelCoordinates && node.excelCoordinates.rowIndex !== undefined) {
      const rowIndex = node.excelCoordinates.rowIndex;
      const excelRow = rowIndex + 1; // Convert to Excel's 1-based row numbering
      
      // Convert column index to Excel column letter (A, B, C, etc.)
      const columnLetter = String.fromCharCode(65 + targetColumnIndex);
      
      // Access the target cell directly 
      const targetRow = rawData[rowIndex];
      
      // Log the entire row to debug issues
      console.log("TARGET ROW:", JSON.stringify(targetRow));
      console.log(`Accessing cell at row ${rowIndex}, column ${targetColumnIndex}`);
      
      // Get the cell value directly
      const value = targetRow && targetColumnIndex < targetRow.length 
        ? targetRow[targetColumnIndex] 
        : null;
      
      const result = {
        rowIndex: rowIndex,
        columnIndex: targetColumnIndex,
        excelRow: excelRow,
        excelColumn: columnLetter,
        excelCell: `${columnLetter}${excelRow}`,
        value: value
      };
      
      console.log(`Excel coordinates: ${result.excelCell}, value:`, value);
      return result;
    }
    
    // Fallback for level 2 nodes (Les) - use the Les number to locate the row
    if (node.columnName === "Les" && node.value) {
      // Extract the Les number (e.g., "Les 1" -> 1)
      const lesMatch = node.value.match(/Les\s+(\d+)/i);
      if (lesMatch && lesMatch[1]) {
        const lesNumber = parseInt(lesMatch[1]);
        
        // Simple mapping: Les 1 -> row index 1 (row 2 in Excel)
        const rowIndex = lesNumber;
        
        console.log(`Using Les number to find row: Les ${lesNumber} -> row index ${rowIndex}`);
        
        if (rowIndex >= 0 && rowIndex < rawData.length) {
          const targetRow = rawData[rowIndex];
          const excelRow = rowIndex + 1;
          const columnLetter = String.fromCharCode(65 + targetColumnIndex);
          
          // Log the row we found
          console.log("FOUND ROW:", JSON.stringify(targetRow));
          
          // Get the cell value directly
          const value = targetRow && targetColumnIndex < targetRow.length 
            ? targetRow[targetColumnIndex] 
            : null;
          
          const result = {
            rowIndex: rowIndex,
            columnIndex: targetColumnIndex,
            excelRow: excelRow,
            excelColumn: columnLetter,
            excelCell: `${columnLetter}${excelRow}`,
            value: value
          };
          
          console.log(`Excel coordinates by Les number: ${result.excelCell}, value:`, value);
          return result;
        }
      }
    }
    
    // Last resort: search for the row with matching values in key columns
    if (node.columnName && node.value) {
      // Determine which column this node's value should be in
      let columnToSearch = -1;
      if (node.columnName === "Blok") columnToSearch = 0;
      else if (node.columnName === "Week") columnToSearch = 1;
      else if (node.columnName === "Les") columnToSearch = 2;
      
      if (columnToSearch >= 0) {
        console.log(`Trying to find '${node.value}' in column ${columnToSearch}`);
        
        // Search for row with matching value
        for (let i = 1; i < rawData.length; i++) {
          if (rawData[i][columnToSearch] === node.value) {
            console.log(`Found matching row at index ${i} for ${node.value}`);
            
            const rowIndex = i;
            const excelRow = rowIndex + 1;
            const columnLetter = String.fromCharCode(65 + targetColumnIndex);
            
            // Get the cell value directly
            const value = rawData[i] && targetColumnIndex < rawData[i].length 
              ? rawData[i][targetColumnIndex] 
              : null;
            
            const result = {
              rowIndex: rowIndex,
              columnIndex: targetColumnIndex,
              excelRow: excelRow,
              excelColumn: columnLetter,
              excelCell: `${columnLetter}${excelRow}`,
              value: value
            };
            
            console.log(`Excel coordinates by search: ${result.excelCell}, value:`, value);
            return result;
          }
        }
      }
    }
    
    console.error("No Excel coordinates could be determined for node:", node);
    return null;
  }
  
  // Get cell value directly using coordinates
  function getExcelCellValue(coordinates) {
    if (!coordinates) return null;
    
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    const rawData = sheetData?.data || [];
    
    if (!rawData || coordinates.rowIndex >= rawData.length) {
      return null;
    }
    
    const row = rawData[coordinates.rowIndex];
    if (!row || coordinates.columnIndex >= row.length) {
      return null;
    }
    
    return row[coordinates.columnIndex];
  }
  
  function addElementDirectly(columnIndex, elementName) {
    // Get Excel data for the active sheet
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    
    if (!sheetData) {
      alert("No sheet data loaded");
      return;
    }
    
    const rawData = sheetData.data || [];
    const headers = sheetData.headers || [];
    
    // Make sure we have data
    if (rawData.length <= 1) {
      alert("No data found in Excel file");
      return;
    }
    
    // Find the active node (current parent in the hierarchy)
    const activeNode = sheetData.root.children[0];
    if (!activeNode) {
      alert("Cannot find active top-level node");
      return;
    }
    
    // Use our new utility function to get exact Excel cell coordinates
    const coordinates = getExcelCoordinates(activeNode, columnIndex);
    
    if (!coordinates) {
      alert("Could not determine Excel coordinates for this element");
      return;
    }
    
    console.log(`Using Excel cell ${coordinates.excelCell} for ${elementName}`);
    
    // Get the value from the specified cell
    const value = coordinates.value;
    
    if (!value) {
      alert(`No value found in Excel cell ${coordinates.excelCell}`);
      return;
    }
    
    console.log(`Found value "${value}" in Excel cell ${coordinates.excelCell}`);
    
    // Check if this element already exists
    if (activeNode.children) {
      const existing = activeNode.children.find(child => 
        child.columnIndex === columnIndex && child.value === value
      );
      
      if (existing) {
        alert(`An element for "${elementName}" with value "${value}" already exists`);
        return;
      }
    }
    
    // Create a new node
    const newNode = {
      id: `node-${Date.now()}-${Math.floor(Math.random() * 10000)}`,
      columnName: elementName,
      columnIndex: columnIndex,
      value: value,
      level: activeNode.level + 1,
      children: [],
      properties: []
    };
    
    // Add properties from other columns in the same row
    const targetRow = rawData[coordinates.rowIndex];
    for (let i = 0; i < targetRow.length; i++) {
      // Skip hierarchy columns and the target column
      if (i === columnIndex || i === activeNode.columnIndex) continue;
      
      const propValue = targetRow[i];
      if (propValue && propValue.toString().trim() !== '') {
        newNode.properties.push({
          columnIndex: i,
          columnName: headers[i] || `Column ${i+1}`,
          value: propValue
        });
      }
    }
    
    // Add to parent's children
    if (!activeNode.children) {
      activeNode.children = [];
    }
    
    activeNode.children.push(newNode);
    
    // Re-render sheet
    renderSheet(activeSheetId, sheetData);
  }
  
  // Helper function to get saved hierarchy configuration
  function getSavedHierarchyConfiguration(fileId, sheetId) {
    // Create a unique key for this file/sheet combination
    const configKey = `${fileId}-${sheetId}`;
    
    // Try to load from in-memory saved configuration first
    let savedConfig = window.excelData.savedConfigurations[configKey];
    
    // If not found in memory but localStorage has configurations, reload from localStorage
    if (!savedConfig) {
      try {
        const savedHierarchyConfigs = localStorage.getItem('hierarchyConfigurations');
        if (savedHierarchyConfigs) {
          const allConfigs = JSON.parse(savedHierarchyConfigs);
          if (allConfigs && typeof allConfigs === 'object' && allConfigs[configKey]) {
            savedConfig = allConfigs[configKey];
            console.log('Found saved configuration in localStorage:', savedConfig);
            
            // Update in-memory version too
            window.excelData.savedConfigurations[configKey] = savedConfig;
          }
        }
      } catch (e) {
        console.error('Error loading hierarchy configuration from localStorage:', e);
      }
    }
    
    return savedConfig;
  }
  
  // Helper function to find a node's parent in the hierarchy
  function findParentNode(nodeId) {
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    
    if (!sheetData || !sheetData.root) return null;
    
    // Recursive function to search the hierarchy
    function searchChildren(parent) {
      if (!parent.children) return null;
      
      // Check if any direct children match the target node ID
      for (const child of parent.children) {
        if (child.id === nodeId) {
          return parent;
        }
        
        // Search deeper in the hierarchy
        const found = searchChildren(child);
        if (found) return found;
      }
      
      return null;
    }
    
    return searchChildren(sheetData.root);
  }
  
  // Function to store current expand/collapse state
  function storeExpandState() {
    const expandStateMap = {};
    
    // Get all nodes in the DOM
    const nodes = document.querySelectorAll('[data-node-id]');
    console.log(`Stored expand state for ${nodes.length} nodes`);
    
    nodes.forEach(node => {
      // Get the node ID and level
      const nodeId = node.getAttribute('data-node-id');
      const level = parseInt(node.getAttribute('data-level') || '0');
      
      if (nodeId) {
        // Store whether the node is expanded (not collapsed)
        expandStateMap[nodeId] = !node.classList.contains(`level-${level}-collapsed`);
      }
    });
    
    return expandStateMap;
  }
  
  // Function to add a new child node and preserve expand/collapse state
  function addNode(parentNode, newNode) {
    if (!parentNode.children) {
      parentNode.children = [];
    }
    
    // Store expand state before modifying the DOM
    storeExpandState();
    
    // Add the new node to the parent
    parentNode.children.push(newNode);
    
    // Re-render with preserved expand state
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    renderSheet(activeSheetId, sheetData);
  }
  
  function deleteNode(nodeId, skipRefresh = true) {
    console.log(`Deleting node ${nodeId}, skipRefresh: ${skipRefresh}`);
    const activeSheetId = window.excelData.activeSheetId;
    const rootNode = window.excelData.sheetsLoaded[activeSheetId].rootNode;
    
    // Store current expand state before deleting
    const expandStateMap = skipRefresh ? storeExpandState() : null;
    
    // Find the node's parent by recursively searching
    let found = false;
    
    function searchAndRemove(parent) {
      if (!parent || !parent.children) return false;
      
      // Check direct children
      for (let i = 0; i < parent.children.length; i++) {
        const child = parent.children[i];
        if (child.id === nodeId) {
          // Found it - remove this node
          parent.children.splice(i, 1);
          console.log(`Node ${nodeId} removed from parent`);
          found = true;
          return true;
        }
      }
      
      // Not found in direct children, check descendants
      for (let i = 0; i < parent.children.length; i++) {
        if (searchAndRemove(parent.children[i])) {
          return true;
        }
      }
      
      return false;
    }
    
    // Start search from root
    searchAndRemove(rootNode);
    
    if (!found) {
      console.error(`Node ${nodeId} not found for deletion`);
      return false;
    }
    
    // If we need to refresh, re-render with preserved expand state
    if (!skipRefresh) {
      console.log('Skipping UI refresh after node deletion');
      return true;
    }
    
    console.log('Re-rendering UI after node deletion');
    // Get the container
    const containerEl = document.getElementById('data-container');
    if (containerEl) {
      // Clear and re-render with the stored expand state
      containerEl.innerHTML = '';
      renderNode(rootNode, 0, expandStateMap);
    }
    
    return true;
  }
  
  // Show modal to add a new Les to a Week node
  function showAddLesModal(weekNode) {
    console.log("showAddLesModal - Raw node data:", weekNode);

    // Get all existing Les nodes for this week
    const existingLesValues = [];
    if (weekNode.children) {
      weekNode.children.forEach(child => {
        if (child.columnName === 'Les') {
          existingLesValues.push(child.value);
        }
      });
    }
    
    console.log("Initial existing Les values:", existingLesValues);
    
    // Create modal container
    const modal = document.createElement('div');
    modal.className = 'element-modal';
    modal.style.position = 'fixed';
    modal.style.top = '0';
    modal.style.left = '0';
    modal.style.width = '100%';
    modal.style.height = '100%';
    modal.style.backgroundColor = 'rgba(0, 0, 0, 0.5)';
    modal.style.display = 'flex';
    modal.style.justifyContent = 'center';
    modal.style.alignItems = 'center';
    modal.style.zIndex = '1000';
    
    // Create content container
    const content = document.createElement('div');
    content.className = 'element-modal-content';
    content.style.backgroundColor = 'white';
    content.style.padding = '20px';
    content.style.borderRadius = '8px';
    content.style.maxWidth = '500px';
    content.style.width = '90%';
    
    // Create header
    const header = document.createElement('h2');
    header.textContent = `Add Les to: ${weekNode.columnName} (${weekNode.value})`;
    header.style.marginBottom = '15px';
    content.appendChild(header);
    
    // Create form
    const form = document.createElement('form');
    form.style.display = 'flex';
    form.style.flexDirection = 'column';
    form.style.gap = '15px';
    
    // Get Excel data
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    
    // Get the raw data directly from the API response or window storage
    let rawData = [];
    if (window.rawExcelData && window.rawExcelData[activeSheetId]) {
      rawData = window.rawExcelData[activeSheetId];
      console.log("Using raw Excel data from window.rawExcelData, length:", rawData.length);
    } else if (sheetData?.data?.data && Array.isArray(sheetData.data.data)) {
      rawData = sheetData.data.data;
      console.log("Using data from sheetData.data.data, length:", rawData.length);
    }
    
    // Les selector
    const lesSelectGroup = document.createElement('div');
    const lesSelectLabel = document.createElement('label');
    lesSelectLabel.textContent = 'Select Les:';
    lesSelectLabel.style.fontWeight = 'bold';
    lesSelectLabel.style.marginBottom = '5px';
    lesSelectLabel.style.display = 'block';
    
    const lesSelect = document.createElement('select');
    lesSelect.style.width = '100%';
    lesSelect.style.padding = '8px';
    lesSelect.style.borderRadius = '4px';
    lesSelect.style.border = '1px solid #ced4da';
    
    // Find all available Les nodes for this week from the raw data
    const availableLesOptions = [];
    const weekValue = weekNode.value.trim();

    // Find which Blok this Week belongs to by traversing the hierarchy
    let blokNode = null;
    let currentNode = weekNode;
    while (currentNode.parent) {
      currentNode = currentNode.parent;
      if (currentNode.columnName === 'Blok') {
        blokNode = currentNode;
        break;
      }
    }
    
    const blokValue = blokNode ? blokNode.value.trim() : "";
    console.log(`Looking for Les nodes for Blok node:`, blokNode);
    console.log(`Looking for Les nodes for Blok: "${blokValue}", Week: "${weekValue}"`);
    
    // Track all possible Les nodes from the Excel data for this Week
    const allLesPossibilities = [];
    
    // DEBUG: Print out a sample of the raw data
    if (rawData && rawData.length > 0) {
      console.log("Excel headers:", rawData[0]);
      if (rawData.length > 3) {
        console.log("Sample data row 1:", rawData[1]);
        console.log("Sample data row 2:", rawData[2]);
        console.log("Sample data row 3:", rawData[3]);
      }
    }
    
    if (!rawData || rawData.length <= 1) {
      console.error("No raw data available to process Les nodes");
    } else {
      console.log("Processing Excel data with", rawData.length, "rows");
      
      // Get the Excel column indices
      const headers = rawData[0];
      const blokColIndex = 0;
      const weekColIndex = 1;
      const lesColIndex = 2;
      
      // Track the context for empty cells
      let currentContextBlok = "";
      let currentContextWeek = "";
      
      // First, directly extract all Les nodes from every row in the raw data
      const allLesNodesInExcel = [];
      
      for (let i = 1; i < rawData.length; i++) {
        const row = rawData[i];
        if (!row || row.length < 3) continue;
        
        // Update context when non-empty cells are found
        if (row[blokColIndex] && row[blokColIndex].trim() !== "") {
          currentContextBlok = row[blokColIndex].trim();
        }
        
        if (row[weekColIndex] && row[weekColIndex].trim() !== "") {
          currentContextWeek = row[weekColIndex].trim();
        }
        
        const rowLes = row[lesColIndex] ? row[lesColIndex].trim() : "";
        if (!rowLes) continue;
        
        // Add to the list of all Les nodes with their context
        allLesNodesInExcel.push({
          blok: currentContextBlok,
          week: currentContextWeek,
          les: rowLes,
          rowIndex: i
        });
      }
      
      console.log("All Les nodes found in Excel:", allLesNodesInExcel);
      
      // Now filter for only the ones matching our Week/Blok context
      const weekLesNodes = allLesNodesInExcel.filter(node => {
        // Direct comparison with our Week node
        const weekMatches = node.week.toLowerCase() === weekValue.toLowerCase();
        
        // If we have a Blok node, compare with it too
        let blokMatches = true;
        if (blokValue) {
          blokMatches = node.blok.toLowerCase() === blokValue.toLowerCase();
        }
        
        return weekMatches && blokMatches;
      });
      
      console.log("Les nodes matching current Week/Blok context:", weekLesNodes);
      
      // Add all matching Les nodes to the possibilities list
      weekLesNodes.forEach(node => {
        if (!allLesPossibilities.some(p => p.value.toLowerCase() === node.les.toLowerCase())) {
          allLesPossibilities.push({
            value: node.les,
            rowIndex: node.rowIndex
          });
          
          // If not already in the existing values, add to available options
          if (!existingLesValues.some(v => v.toLowerCase() === node.les.toLowerCase())) {
            availableLesOptions.push({
              value: node.les,
              rowIndex: node.rowIndex
            });
          }
        }
      });
    }
    
    console.log("All Les possibilities for this week:", allLesPossibilities);
    console.log("Existing Les values:", existingLesValues);
    console.log("Available Les options for dropdown:", availableLesOptions);
    
    // Add no options text if none available
    if (availableLesOptions.length === 0) {
      if (allLesPossibilities.length === 0) {
        const noOptionsOption = document.createElement('option');
        noOptionsOption.value = "";
        noOptionsOption.textContent = "No Les nodes found for this Week";
        noOptionsOption.disabled = true;
        lesSelect.appendChild(noOptionsOption);
      } else {
        // All Les nodes for this Week are already added
        const noOptionsOption = document.createElement('option');
        noOptionsOption.value = "";
        noOptionsOption.textContent = "All Les nodes for this Week are already added";
        noOptionsOption.disabled = true;
        lesSelect.appendChild(noOptionsOption);
      }
    } else {
      // Add a default option
      const defaultOption = document.createElement('option');
      defaultOption.value = "";
      defaultOption.textContent = "-- Select a Les --";
      defaultOption.disabled = true;
      defaultOption.selected = true;
      lesSelect.appendChild(defaultOption);
      
      // Add options to select, sorted by Les number if possible
      availableLesOptions.sort((a, b) => {
        // Try to extract numbers from Les names for proper sorting
        const aMatch = a.value.match(/Les\s+(\d+)/i);
        const bMatch = b.value.match(/Les\s+(\d+)/i);
        
        if (aMatch && bMatch) {
          return parseInt(aMatch[1]) - parseInt(bMatch[1]);
        }
        
        // Fallback to alphabetical sorting
        return a.value.localeCompare(b.value);
      }).forEach(option => {
        const lesOption = document.createElement('option');
        lesOption.value = option.value;
        lesOption.textContent = option.value;
        lesOption.dataset.rowIndex = option.rowIndex;
        lesSelect.appendChild(lesOption);
      });
    }
    
    lesSelectGroup.appendChild(lesSelectLabel);
    lesSelectGroup.appendChild(lesSelect);
    form.appendChild(lesSelectGroup);
    
    // Preview section
    const previewContainer = document.createElement('div');
    previewContainer.style.marginTop = '15px';
    previewContainer.style.padding = '10px';
    previewContainer.style.backgroundColor = '#f8f9fa';
    previewContainer.style.borderRadius = '4px';
    previewContainer.style.fontSize = '0.9rem';
    
    const previewHeader = document.createElement('div');
    previewHeader.textContent = 'Preview';
    previewHeader.style.fontWeight = 'bold';
    previewHeader.style.marginBottom = '5px';
    
    const previewContent = document.createElement('div');
    previewContent.style.fontFamily = 'monospace';
    
    previewContainer.appendChild(previewHeader);
    previewContainer.appendChild(previewContent);
    form.appendChild(previewContainer);
    
    // Update preview function
    function updatePreview() {
      const selectedLes = lesSelect.value;
      const selectedOption = lesSelect.options[lesSelect.selectedIndex];
      const rowIndex = selectedOption ? parseInt(selectedOption.dataset.rowIndex) : -1;
      
      if (!selectedLes || rowIndex < 0) {
        previewContent.innerHTML = '<p class="error">No Les selected</p>';
        return;
      }
      
      // Get row data
      const row = rawData[rowIndex];
      
      if (!row) {
        previewContent.innerHTML = '<p class="error">Could not find data for selected Les</p>';
        return;
      }
      
      // Display preview
      previewContent.innerHTML = `
        <p><strong>Les:</strong> ${selectedLes}</p>
        <p><strong>Excel Row:</strong> ${rowIndex + 1}</p>
        <p><strong>Werkblad:</strong> ${row[3] || '(empty)'}</p>
        <p><strong>Instructie & begeleide inoefening - Leerlingen:</strong> ${row[4] || '(empty)'}</p>
      `;
    }
    
    // Update preview when selection changes
    lesSelect.addEventListener('change', updatePreview);
    
    // Add buttons
    const buttonContainer = document.createElement('div');
    buttonContainer.style.display = 'flex';
    buttonContainer.style.justifyContent = 'space-between';
    buttonContainer.style.marginTop = '20px';
    
    const cancelButton = document.createElement('button');
    cancelButton.textContent = 'Cancel';
    cancelButton.type = 'button';
    cancelButton.style.padding = '8px 15px';
    
    const addButton = document.createElement('button');
    addButton.textContent = 'Add Les';
    addButton.type = 'submit';
    addButton.style.padding = '8px 15px';
    addButton.style.backgroundColor = '#34a3d7';
    addButton.style.color = 'white';
    addButton.style.border = 'none';
    
    // Disable add button if no options available
    addButton.disabled = availableLesOptions.length === 0;
    
    buttonContainer.appendChild(cancelButton);
    buttonContainer.appendChild(addButton);
    form.appendChild(buttonContainer);
    
    // Add form to content
    content.appendChild(form);
    
    // Add content to modal
    modal.appendChild(content);
    
    // Add modal to document
    document.body.appendChild(modal);
    
    // Initial preview update
    updatePreview();
    
    // Handle form submission
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      
      const selectedLes = lesSelect.value;
      const selectedOption = lesSelect.options[lesSelect.selectedIndex];
      const rowIndex = selectedOption ? parseInt(selectedOption.dataset.rowIndex) : -1;
      
      console.log("Form submitted with Les:", selectedLes, "row index:", rowIndex);
      
      if (!selectedLes || rowIndex < 0) {
        alert('Please select a Les from the dropdown');
        return;
      }
      
      // Get the Excel row data
      const row = rawData[rowIndex];
      if (!row) {
        alert('Could not find data for the selected Les');
        return;
      }
      
      // Get headers for property names
      const headers = rawData[0] || [];
      
      // Create new Les node
      const newLesNode = {
        id: `node-${Date.now()}-${Math.floor(Math.random() * 10000)}`,
        columnName: 'Les',
        columnIndex: 2, // Index for Les column
        value: selectedLes,
        level: 2,
        children: [],
        properties: [],
        parent: weekNode
      };
      
      // Add Excel coordinates for reference
      newLesNode.excelCoordinates = {
        rowIndex: rowIndex,
        row: rowIndex + 1,
        column: 'C',
        cell: `C${rowIndex + 1}`
      };
      
      // Add properties from the Excel row (all columns after Les)
      for (let i = 3; i < row.length; i++) {
        // Include even empty cells to maintain structure
        newLesNode.properties.push({
          columnIndex: i,
          columnName: headers[i] || `Column ${i+1}`,
          value: row[i] || ''  // Use empty string for null/undefined values
        });
      }
      
      console.log("Adding new Les node:", newLesNode);
      
      // Add the Les node to the Week node
      addNode(weekNode, newLesNode);
      
      // Store the expand state before updating UI
      storeExpandState();
      
      // Close modal
      document.body.removeChild(modal);
      
      // Update the UI without refreshing the page
      const sheetContainer = document.getElementById('sheet-container');
      if (sheetContainer) {
        sheetContainer.innerHTML = '';
        renderNode(window.hierarchicalData, 0);
        applyStyles();
      } else {
        console.error("Could not find sheet-container, UI may not update properly");
      }
    });
    
    // Handle cancel
    cancelButton.addEventListener('click', () => {
      document.body.removeChild(modal);
    });
  }
  


  
  // Helper function to safely store and retrieve raw data
  function storeRawExcelData(sheetId, data) {
    if (!window.rawExcelData) {
      window.rawExcelData = {};
    }
    window.rawExcelData[sheetId] = data;
    console.log(`Stored raw Excel data for sheet ${sheetId} length: ${data.length}`);
  }
  
  // Helper function to get raw Excel data
  function getRawExcelData(sheetId) {
    try {
      // First try getting it from window.rawExcelData
      if (window.rawExcelData && window.rawExcelData[sheetId]) {
        console.log(`Using stored rawExcelData for sheet ${sheetId}`);
        return window.rawExcelData[sheetId];
      }
      
      // Then try getting it from sheetsLoaded if available
      if (window.excelData.sheetsLoaded && 
          window.excelData.sheetsLoaded[sheetId] && 
          window.excelData.sheetsLoaded[sheetId].data && 
          Array.isArray(window.excelData.sheetsLoaded[sheetId].data.data)) {
        console.log(`Using data from sheetsLoaded for sheet ${sheetId}`);
        return window.excelData.sheetsLoaded[sheetId].data.data;
      }
      
      console.warn(`No raw data found for sheet ${sheetId}`);
      return null;
    } catch (error) {
      console.error(`Error retrieving raw Excel data for sheet ${sheetId}:`, error);
      return null;
    }
  }
  
  // Find parent node by traversing the node tree
  function findParentNode(nodeId) {
    const activeSheetId = window.excelData.activeSheetId;
    const rootNode = window.excelData.sheetsLoaded[activeSheetId].rootNode;
    
    // Use a recursive function to search through all nodes
    return searchChildren(rootNode);
    
    function searchChildren(parent) {
      // Skip if this node doesn't have children
      if (!parent || !parent.children || !Array.isArray(parent.children)) {
        return null;
      }
      
      // Check direct children first
      for (const child of parent.children) {
        if (child.id === nodeId) {
          return parent;
        }
      }
      
      // Then recursively check grandchildren
      for (const child of parent.children) {
        const foundParent = searchChildren(child);
        if (foundParent) {
          return foundParent;
        }
      }
      
      return null;
    }
  }
  
  // Add this function after the initToolbar function
  function addParentLevel() {
    console.log("Adding parent level - properly adjusting all hierarchy levels");
    
    // Get active sheet data
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    
    if (!sheetData || !sheetData.headers) {
      alert("No active sheet data found");
      return;
    }

    // Check if this tab already has Course nodes at level -1
    const hasCourseAtLevelMinus1 = sheetData.root.children.some(node => 
      node.columnName === 'Course' && node.level === -1);
    
    if (hasCourseAtLevelMinus1) {
      alert("This tab already has a Course at level -1.");
      return;
    }
    
    // Check if this tab already has Course nodes at level 0
    const courseNodesAtLevel0 = sheetData.root.children.filter(node => 
      node.columnName === 'Course' && node.level === 0);
    
    if (courseNodesAtLevel0.length > 0) {
      // If we have Course nodes at level 0, promote them to level -1
      // and update all child levels recursively
      courseNodesAtLevel0.forEach(courseNode => {
        // Change Course to level -1
        courseNode.level = -1;
        
        // Now update all children recursively using existing function
        updateLevels(courseNode);
      });
      
      console.log("Promoted existing Course nodes to level -1 and adjusted all children");
    } else {
      // If no Course nodes exist, create a new one as parent
      const parentNode = {
        id: `node-course-${Date.now()}`,
        value: "Course",
        columnName: "Course",
        level: -1,
        children: []
      };
      
      // Make a deep copy of existing nodes to avoid reference issues
      const existingNodes = JSON.parse(JSON.stringify(sheetData.root.children));
      
      // Set the existing nodes as children of the Course parent
      parentNode.children = existingNodes;
      
      // Replace root's children with just the Course parent
      sheetData.root.children = [parentNode];
      
      // Update all children recursively to ensure proper level assignment
      updateLevels(parentNode);
      
      console.log("Created new Course parent at level -1 and adjusted all children");
    }
    
    // Add CSS for level-_1 if it doesn't exist yet
    if (!document.getElementById('level-_1-styles')) {
      const style = document.createElement('style');
      style.id = 'level-_1-styles';
      style.textContent = `
        .level-_1-node {
          margin-bottom: 20px;
          border: 1px solid #34a3d7;
          border-radius: 5px;
          overflow: hidden;
        }
        .level-_1-header {
          background-color: #34a3d7;
          color: white;
          padding: 8px 15px;
          display: flex;
          align-items: center;
          cursor: pointer;
          font-weight: bold;
        }
        .level-_1-title {
          flex: 1;
        }
        .level-_1-content {
          padding: 0;
          overflow: hidden;
        }
        .level-_1-children {
          display: flex;
          flex-direction: column;
        }
        .level-_1-expand-button {
          width: 12px;
          height: 12px;
          margin-right: 10px;
          background-image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16"><path fill="white" d="M8 4a.5.5 0 0 1 .5.5v3h3a.5.5 0 0 1 0 1h-3v3a.5.5 0 0 1-1 0v-3h-3a.5.5 0 0 1 0-1h3v-3A.5.5 0 0 1 8 4z"/></svg>');
          background-repeat: no-repeat;
          background-position: center;
          transition: transform 0.2s;
        }
        .level-_1-collapsed .level-_1-expand-button {
          transform: rotate(0deg);
        }
        .level-_1-collapsed > .level-_1-content {
          display: none;
        }
      `;
      document.head.appendChild(style);
    }
    
    // Re-render the sheet with the updated hierarchy
    renderSheet(activeSheetId, sheetData);
    
    console.log("Parent level added successfully with proper hierarchy levels");
  }
  
  // Helper function to ensure CSS styles for all levels are set up
  function ensureLevelStyles() {
    // Add CSS for level-_1 if it doesn't exist yet
    if (!document.getElementById('level-_1-styles')) {
      const style = document.createElement('style');
      style.id = 'level-_1-styles';
      style.textContent = `
        .level-_1-node {
          margin-bottom: 20px;
          border: 1px solid #34a3d7;
          border-radius: 5px;
          overflow: hidden;
        }
        .level-_1-header {
          background-color: #34a3d7;
          color: white;
          padding: 8px 15px;
          display: flex;
          align-items: center;
          cursor: pointer;
          font-weight: bold;
        }
        .level-_1-title {
          flex: 1;
        }
        .level-_1-content {
          padding: 0;
          overflow: hidden;
        }
        .level-_1-children {
          display: flex;
          flex-direction: column;
        }
        .level-_1-expand-button {
          width: 12px;
          height: 12px;
          margin-right: 10px;
          background-image: url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 16 16"><path fill="white" d="M8 4a.5.5 0 0 1 .5.5v3h3a.5.5 0 0 1 0 1h-3v3a.5.5 0 0 1-1 0v-3h-3a.5.5 0 0 1 0-1h3v-3A.5.5 0 0 1 8 4z"/></svg>');
          background-repeat: no-repeat;
          background-position: center;
          transition: transform 0.2s;
        }
        .level-_1-collapsed .level-_1-expand-button {
          transform: rotate(0deg);
        }
        .level-_1-collapsed > .level-_1-content {
          display: none;
        }
      `;
      document.head.appendChild(style);
    }
    
    // Add CSS for level-4 if it doesn't exist yet (for Les children)
    if (!document.getElementById('level-4-styles')) {
      const style = document.createElement('style');
      style.id = 'level-4-styles';
      style.textContent = `
        .level-4-node {
          padding: 8px 10px;
          margin-left: 30px;
          border-left: 3px solid #17a2b8;
          background-color: #f8f9fa;
          margin-bottom: 5px;
        }
        .level-4-header {
          display: flex;
          align-items: center;
          font-size: 0.9em;
        }
        .level-4-title {
          flex: 1;
          margin-left: 10px;
        }
      `;
      document.head.appendChild(style);
    }
  }
  
  // Helper function to update node levels recursively
  function updateLevels(node) {
    if (node.children && node.children.length > 0) {
      node.children.forEach(child => {
        // Child level is parent level + 1
        child.level = node.level + 1;
        // Recursively update grandchildren
        updateLevels(child);
      });
    }
  }
  
  // Function to add a child container
  function addChildContainer(parentNode) {
    const newNode = {
      id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
      value: 'New Container',
      columnName: '',
      level: parentNode.level + 1,
      children: [],
      properties: []
    };
    
    if (!parentNode.children) {
      parentNode.children = [];
    }
    
    parentNode.children.push(newNode);
    
    // Re-render
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    renderSheet(activeSheetId, sheetData);
  }
  
  // Function to add a sibling container
  function addSiblingContainer(node) {
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    
    // Find parent
    const parent = findParentNode(node.id) || sheetData.root;
    
    const newNode = {
      id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
      value: 'New Container',
      columnName: '',
      level: node.level,
      children: [],
      properties: []
    };
    
    if (!parent.children) {
      parent.children = [];
    }
    
    // Insert after current node
    const currentIndex = parent.children.findIndex(child => child.id === node.id);
    parent.children.splice(currentIndex + 1, 0, newNode);
    
    // Re-render
    renderSheet(activeSheetId, sheetData);
  }
  
  // Function to delete a container
  function deleteContainer(node) {
    if (confirm(`Delete container "${node.value || node.columnName || 'New Container'}" and all its children?`)) {
      const activeSheetId = window.excelData.activeSheetId;
      const sheetData = window.excelData.sheetsLoaded[activeSheetId];
      
      // Find parent and remove node
      const parent = findParentNode(node.id) || sheetData.root;
      if (parent && parent.children) {
        parent.children = parent.children.filter(child => child.id !== node.id);
      }
      
      // Re-render
      renderSheet(activeSheetId, sheetData);
    }
  }
  

  
  // Function to initialize toolbar
  function initToolbar() {
    // Add event listeners for collapse/expand all buttons
    if (collapseAllButton) {
      collapseAllButton.addEventListener('click', () => toggleAllNodes(true));
    }
    if (expandAllButton) {
      expandAllButton.addEventListener('click', () => toggleAllNodes(false));
    }
    
    // Add event listener for build structure button
    const buildStructureButton = document.getElementById('build-structure');
    if (buildStructureButton) {
      buildStructureButton.addEventListener('click', showStructureBuilder);
    }
    
    // Add Parent Level button
    const addParentLevelButton = document.createElement('button');
    addParentLevelButton.textContent = 'Add Parent Level';
    addParentLevelButton.className = 'toolbar-button';
    addParentLevelButton.style.marginLeft = '10px';
    
    addParentLevelButton.addEventListener('click', addParentLevel);
    
    // Add to toolbar after other buttons
    const toolbar = document.querySelector('.toolbar');
    if (toolbar) {
      toolbar.appendChild(addParentLevelButton);
    }
  }
  
  // Function to show the structure builder modal
  function showStructureBuilder() {
    const activeSheetId = window.excelData.activeSheetId;
    if (!activeSheetId) {
      alert('Please select a sheet first');
      return;
    }
    
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    if (!sheetData || !sheetData.headers) {
      alert('No data available for this sheet');
      return;
    }
    
    // Create modal
    const modal = document.createElement('div');
    modal.className = 'pdf-preview-modal';
    
    // Create content container
    const content = document.createElement('div');
    content.className = 'pdf-preview-content';
    content.style.maxWidth = '900px';
    content.style.width = '90%';
    content.style.height = '80vh';
    content.style.display = 'flex';
    content.style.flexDirection = 'column';
    
    // Add close button
    const closeButton = document.createElement('button');
    closeButton.className = 'pdf-preview-close';
    closeButton.innerHTML = '&times;';
    closeButton.addEventListener('click', () => {
      document.body.removeChild(modal);
    });
    
    // Add header
    const header = document.createElement('h2');
    header.textContent = 'Build Custom Structure';
    header.style.marginBottom = '20px';
    header.style.color = '#34a3d7';
    
    // Create main container with two panels
    const mainContainer = document.createElement('div');
    mainContainer.style.display = 'flex';
    mainContainer.style.gap = '20px';
    mainContainer.style.flex = '1';
    mainContainer.style.overflow = 'hidden';
    
    // Left panel - Structure builder
    const leftPanel = document.createElement('div');
    leftPanel.style.flex = '1';
    leftPanel.style.display = 'flex';
    leftPanel.style.flexDirection = 'column';
    leftPanel.style.borderRight = '1px solid #ddd';
    leftPanel.style.paddingRight = '20px';
    
    const structureTitle = document.createElement('h3');
    structureTitle.textContent = 'Structure';
    structureTitle.style.marginBottom = '10px';
    
    const structureContainer = document.createElement('div');
    structureContainer.id = 'custom-structure-container';
    structureContainer.style.flex = '1';
    structureContainer.style.overflow = 'auto';
    structureContainer.style.border = '1px solid #ddd';
    structureContainer.style.borderRadius = '4px';
    structureContainer.style.padding = '10px';
    structureContainer.style.backgroundColor = '#f9f9f9';
    
    // Add root container
    const rootContainer = createStructureNode('Root', null, true);
    structureContainer.appendChild(rootContainer);
    
    leftPanel.appendChild(structureTitle);
    leftPanel.appendChild(structureContainer);
    
    // Right panel - Excel columns
    const rightPanel = document.createElement('div');
    rightPanel.style.flex = '1';
    rightPanel.style.display = 'flex';
    rightPanel.style.flexDirection = 'column';
    
    const columnsTitle = document.createElement('h3');
    columnsTitle.textContent = 'Available Excel Columns';
    columnsTitle.style.marginBottom = '10px';
    
    const columnsContainer = document.createElement('div');
    columnsContainer.style.flex = '1';
    columnsContainer.style.overflow = 'auto';
    columnsContainer.style.border = '1px solid #ddd';
    columnsContainer.style.borderRadius = '4px';
    columnsContainer.style.padding = '10px';
    columnsContainer.style.backgroundColor = '#f9f9f9';
    
    // Add columns from Excel
    sheetData.headers.forEach((header, index) => {
      if (!header) return;
      
      const columnItem = document.createElement('div');
      columnItem.className = 'draggable-column';
      columnItem.draggable = true;
      columnItem.dataset.columnIndex = index;
      columnItem.dataset.columnName = header;
      columnItem.textContent = `${header} (Column ${index + 1})`;
      columnItem.style.padding = '8px';
      columnItem.style.margin = '5px 0';
      columnItem.style.backgroundColor = '#e3f2fd';
      columnItem.style.borderRadius = '4px';
      columnItem.style.cursor = 'move';
      columnItem.style.border = '1px solid #90caf9';
      
      // Add drag events
      columnItem.addEventListener('dragstart', (e) => {
        e.dataTransfer.setData('columnIndex', index);
        e.dataTransfer.setData('columnName', header);
        columnItem.style.opacity = '0.5';
      });
      
      columnItem.addEventListener('dragend', () => {
        columnItem.style.opacity = '1';
      });
      
      columnsContainer.appendChild(columnItem);
    });
    
    rightPanel.appendChild(columnsTitle);
    rightPanel.appendChild(columnsContainer);
    
    // Add panels to main container
    mainContainer.appendChild(leftPanel);
    mainContainer.appendChild(rightPanel);
    
    // Create buttons
    const buttonContainer = document.createElement('div');
    buttonContainer.style.display = 'flex';
    buttonContainer.style.justifyContent = 'space-between';
    buttonContainer.style.marginTop = '20px';
    buttonContainer.style.paddingTop = '20px';
    buttonContainer.style.borderTop = '1px solid #ddd';
    
    const cancelButton = document.createElement('button');
    cancelButton.textContent = 'Cancel';
    cancelButton.className = 'control-button';
    
    const applyButton = document.createElement('button');
    applyButton.textContent = 'Apply Structure';
    applyButton.className = 'control-button';
    applyButton.style.backgroundColor = '#34a3d7';
    applyButton.style.color = 'white';
    
    buttonContainer.appendChild(cancelButton);
    buttonContainer.appendChild(applyButton);
    
    // Assemble modal
    content.appendChild(closeButton);
    content.appendChild(header);
    content.appendChild(mainContainer);
    content.appendChild(buttonContainer);
    modal.appendChild(content);
    
    // Add to document
    document.body.appendChild(modal);
    
    // Event handlers
    cancelButton.addEventListener('click', () => {
      document.body.removeChild(modal);
    });
    
    applyButton.addEventListener('click', () => {
      const structure = extractStructure(rootContainer);
      applyCustomStructure(structure);
      document.body.removeChild(modal);
    });
  }
  
  // Helper function to create a structure node
  function createStructureNode(name, columnInfo = null, isRoot = false) {
    const node = document.createElement('div');
    node.className = 'structure-node';
    node.style.marginBottom = '10px';
    
    const header = document.createElement('div');
    header.className = 'structure-node-header';
    header.style.display = 'flex';
    header.style.alignItems = 'center';
    header.style.padding = '8px';
    header.style.backgroundColor = isRoot ? '#34a3d7' : '#e0e0e0';
    header.style.color = isRoot ? 'white' : 'black';
    header.style.borderRadius = '4px';
    header.style.marginBottom = '5px';
    
    const title = document.createElement('span');
    title.textContent = name;
    title.style.flex = '1';
    
    // Column selector
    const columnSelect = document.createElement('select');
    columnSelect.style.marginLeft = '10px';
    columnSelect.style.padding = '4px';
    columnSelect.style.borderRadius = '4px';
    
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = 'No column';
    columnSelect.appendChild(defaultOption);
    
    // Add Excel columns as options
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    if (sheetData && sheetData.headers) {
      sheetData.headers.forEach((header, index) => {
        if (!header) return;
        const option = document.createElement('option');
        option.value = index;
        option.textContent = header;
        columnSelect.appendChild(option);
      });
    }
    
    if (columnInfo) {
      columnSelect.value = columnInfo.columnIndex;
    }
    
    // Add container button
    const addButton = document.createElement('button');
    addButton.textContent = '+ Add Container';
    addButton.style.marginLeft = '10px';
    addButton.style.padding = '4px 8px';
    addButton.style.fontSize = '12px';
    addButton.style.cursor = 'pointer';
    
    // Delete button (not for root)
    if (!isRoot) {
      const deleteButton = document.createElement('button');
      deleteButton.textContent = '×';
      deleteButton.style.marginLeft = '10px';
      deleteButton.style.padding = '4px 8px';
      deleteButton.style.fontSize = '16px';
      deleteButton.style.cursor = 'pointer';
      deleteButton.style.backgroundColor = '#f44336';
      deleteButton.style.color = 'white';
      deleteButton.style.border = 'none';
      deleteButton.style.borderRadius = '4px';
      
      deleteButton.addEventListener('click', () => {
        node.remove();
      });
      
      header.appendChild(deleteButton);
    }
    
    header.appendChild(title);
    header.appendChild(columnSelect);
    header.appendChild(addButton);
    
    const childrenContainer = document.createElement('div');
    childrenContainer.className = 'structure-node-children';
    childrenContainer.style.marginLeft = '20px';
    childrenContainer.style.paddingLeft = '10px';
    childrenContainer.style.borderLeft = '2px solid #ddd';
    
    // Make it a drop zone
    childrenContainer.addEventListener('dragover', (e) => {
      e.preventDefault();
      childrenContainer.style.backgroundColor = '#e8f5e9';
    });
    
    childrenContainer.addEventListener('dragleave', () => {
      childrenContainer.style.backgroundColor = 'transparent';
    });
    
    childrenContainer.addEventListener('drop', (e) => {
      e.preventDefault();
      childrenContainer.style.backgroundColor = 'transparent';
      
      const columnIndex = e.dataTransfer.getData('columnIndex');
      const columnName = e.dataTransfer.getData('columnName');
      
      if (columnIndex && columnName) {
        const newNode = createStructureNode(`Container (${columnName})`, {
          columnIndex: parseInt(columnIndex),
          columnName: columnName
        });
        childrenContainer.appendChild(newNode);
      }
    });
    
    node.appendChild(header);
    node.appendChild(childrenContainer);
    
    // Add container button handler
    addButton.addEventListener('click', () => {
      const newNode = createStructureNode('New Container');
      childrenContainer.appendChild(newNode);
    });
    
    // Store column info
    node.dataset.columnIndex = columnInfo ? columnInfo.columnIndex : '';
    node.dataset.columnName = columnInfo ? columnInfo.columnName : '';
    
    return node;
  }
  
  // Helper function to extract structure from DOM
  function extractStructure(node) {
    const columnSelect = node.querySelector('.structure-node-header select');
    const childrenContainer = node.querySelector('.structure-node-children');
    const children = [];
    
    if (childrenContainer) {
      const childNodes = childrenContainer.querySelectorAll(':scope > .structure-node');
      childNodes.forEach(childNode => {
        children.push(extractStructure(childNode));
      });
    }
    
    return {
      columnIndex: columnSelect ? parseInt(columnSelect.value) : -1,
      columnName: columnSelect && columnSelect.value ? 
        columnSelect.options[columnSelect.selectedIndex].text : '',
      children: children
    };
  }
  
  // Function to apply custom structure to the data
  function applyCustomStructure(structure) {
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    
    if (!sheetData) return;
    
    // Get raw Excel data
    const rawData = getRawExcelData(activeSheetId);
    if (!rawData) {
      alert('No raw data available');
      return;
    }
    
    // Build new hierarchical structure based on custom structure
    const newRoot = {
      type: 'root',
      children: [],
      level: -1
    };
    
    // Process each row of Excel data
    rawData.slice(1).forEach((row, rowIndex) => {
      if (row.every(cell => !cell)) return; // Skip empty rows
      
      // Create nodes based on structure
      processStructureRow(structure, row, newRoot, 0, rowIndex + 1);
    });
    
    // Update sheet data with new structure
    sheetData.root = newRoot;
    sheetData.customStructure = structure;
    
    // Re-render the sheet
    renderSheet(activeSheetId, sheetData);
  }
  
  // Helper function to process a row based on custom structure
  function processStructureRow(structure, row, parentNode, level, rowIndex) {
    if (structure.columnIndex === -1 || !structure.columnName) {
      // This is just a container without data
      // Process children with the same parent
      structure.children.forEach(childStructure => {
        processStructureRow(childStructure, row, parentNode, level, rowIndex);
      });
      return;
    }
    
    const cellValue = row[structure.columnIndex] || '';
    if (!cellValue && structure.children.length === 0) return; // Skip empty cells with no children
    
    // Find or create node
    let node = parentNode.children.find(child => 
      child.value === cellValue && 
      child.columnIndex === structure.columnIndex
    );
    
    if (!node) {
      node = {
        id: `node-${structure.columnIndex}-${cellValue}-${Date.now()}-${Math.random()}`,
        type: 'item',
        value: cellValue,
        columnName: structure.columnName,
        columnIndex: structure.columnIndex,
        level: level,
        children: [],
        properties: [],
        rowIndex: rowIndex
      };
      parentNode.children.push(node);
    }
    
    // Process children
    structure.children.forEach(childStructure => {
      processStructureRow(childStructure, row, node, level + 1, rowIndex);
    });
  }
  
  initialize();
  
  // Export functions for use by live-editor.js
  window.renderSheet = renderSheet;
}); 