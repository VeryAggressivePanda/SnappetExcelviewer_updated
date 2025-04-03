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
    
    // Get level-1 node values from active sheet data
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    if (!sheetData || !sheetData.root || !sheetData.root.children) {
      return;
    }
    
    // Get only level-1 nodes (first column)
    const level1Nodes = sheetData.root.children.filter(node => node.columnIndex === 0);
    
    // Add default option
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.text = 'Select an item to export...';
    exportColumnSelect.appendChild(defaultOption);
    
    // Add option for each level-1 node
    level1Nodes.forEach(node => {
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
        alert('Please select a level-1 item to export');
        return;
      }
      
      // Get the data for the active sheet
      const sheetData = window.excelData.sheetsLoaded[activeSheetId];
      
      // Filter the data by the selected level-1 value
      let filteredData = { ...sheetData };
      if (selectedValue) {
        // Only include the selected level-1 node and its children
        filteredData = {
          ...sheetData,
          root: {
            ...sheetData.root,
            children: sheetData.root.children.filter(node => 
              node.value === selectedValue && node.columnIndex === 0
            )
          }
        };
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
        alert('Please select a level-1 item to preview');
        return;
      }
      
      // Get the data for the active sheet
      const sheetData = window.excelData.sheetsLoaded[activeSheetId];
      
      // Filter the data by the selected level-1 value
      let filteredData = { ...sheetData };
      if (selectedValue) {
        // Only include the selected level-1 node and its children
        filteredData = {
          ...sheetData,
          root: {
            ...sheetData.root,
            children: sheetData.root.children.filter(node => 
              node.value === selectedValue && node.columnIndex === 0
            )
          }
        };
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
      exportPdf();
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
    
    // Process all rows
    for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
      const row = data[rowIndex];
      
      // Skip empty rows
      if (row.every(cell => !cell)) continue;
      
      // Update current values for each column
      for (const colIndex in hierarchy) {
        const i = parseInt(colIndex);
        if (row[i] && row[i].trim() !== '') {
          currentValues[i] = row[i];
        }
      }
      
      // Process each top-level column
      for (const topColIndex of topLevelColumns) {
        // Use current value for this column (may come from a previous row)
        const topValue = currentValues[topColIndex] || '';
        if (!topValue) continue;
        
        // Create or get top-level node - use value-based key, not row-based
        const topKey = `col${topColIndex}-${topValue}`;
        if (!nodesByColumn[topKey]) {
          const node = {
            id: `node-${topColIndex}-${topValue.replace(/[^a-zA-Z0-9]/g, '-')}`,
            value: topValue,
            columnName: headers[topColIndex] || `Column ${topColIndex+1}`,
            columnIndex: topColIndex,
            level: 0,
            children: [],
            properties: [],
            childNodes: {} // Temporary map to track child nodes
          };
          root.children.push(node);
          nodesByColumn[topKey] = node;
        }
        
        // Process child columns recursively using current values
        processChildrenWithCurrentValues(
          nodesByColumn[topKey], 
          row, 
          0, 
          topColIndex, 
          rowIndex, 
          childrenMap, 
          headers, 
          hierarchy, 
          currentValues
        );
      }
    }
    
    // Clean up temporary tracking objects
    for (const key in nodesByColumn) {
      cleanupNode(nodesByColumn[key]);
    }
    
    // Log the result
    console.log("Processed data structure:", {
      topLevelCount: root.children.length,
      levels: countLevels(root)
    });
    
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
      
      // Create a unique key for this child based on column and value
      // Remove rowIndex from key to ensure grouping by value
      const childKey = `${parentNode.id}-col${childColIndex}-${childValue || 'empty'}`;
      let childNode = parentNode.childNodes[childKey];
      
      if (!childNode) {
        childNode = {
          id: `node-${childColIndex}-${childValue.replace(/[^a-zA-Z0-9]/g, '-')}`, // Make ID based on value, not row
          value: childValue || `Empty ${headers[childColIndex]}`,
          columnName: headers[childColIndex] || `Column ${childColIndex+1}`,
          columnIndex: childColIndex,
          level: level + 1,
          children: [],
          properties: [],
          childNodes: {}, // For its own children
          isEmpty: !childValue // Track if this is an empty cell
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
      
      // Add property if not already exists with this column index
      const existingProp = node.properties.find(p => p.columnIndex === i);
      if (!existingProp) {
        node.properties.push({
          columnIndex: i,
          columnName: columnName,
          value: value
        });
      }
    }
  }
  
  // Helper to count nodes at each level
  function countLevels(node) {
    const counts = {0: 0, 1: 0, 2: 0};
    
    function traverse(node, level) {
      if (level in counts) {
        counts[level]++;
      }
      
      if (node.children) {
        for (const child of node.children) {
          traverse(child, level + 1);
        }
      }
    }
    
    for (const child of node.children) {
      traverse(child, 0);
    }
    
    return counts;
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
  
  // Load sheet data via AJAX
  async function loadSheetData(sheetId) {
    try {
      const sheetContent = document.getElementById(`sheet-${sheetId}`);
      
      // Show loading indicator
      sheetContent.innerHTML = '<div class="loading-indicator"><p>Loading sheet data...</p></div>';
      
      console.log(`Loading sheet data for sheet ${sheetId}, using fileId: ${window.excelData.fileId}`);
      
      if (!window.excelData.fileId) {
        throw new Error('No file ID available for API request');
      }
      
      // Fetch the sheet data
      const url = `/api/sheet/${window.excelData.fileId}/${sheetId}`;
      console.log(`Requesting: ${url}`);
      
      const response = await fetch(url);
      
      if (!response.ok) {
        const errorText = await response.text();
        console.error(`API response error: ${response.status} ${response.statusText}`, errorText);
        throw new Error(`Failed to load sheet data: ${response.statusText}`);
      }
      
      console.log('Response received, parsing JSON...');
      const result = await response.json();
      console.log('API response parsed:', result);
      
      if (!result.data) {
        throw new Error('Invalid response from server');
      }
      
      const data = result.data;
      
      // If we have raw data, check for saved configuration first
      if (data.needsConfiguration && data.headers && data.data) {
        // Create a unique key for this file/sheet combination
        const configKey = `${window.excelData.fileId}-${sheetId}`;
        
        // Try to load from in-memory saved configuration first
        let savedConfig = window.excelData.savedConfigurations[configKey];
        
        // If not found in memory but localStorage has configurations, reload from localStorage
        // This handles page refresh case when memory is cleared but localStorage persists
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
        
        if (savedConfig) {
          console.log('Using saved hierarchy configuration:', savedConfig);
          
          // Store hierarchy configuration
          window.excelData.hierarchyConfigurations[sheetId] = savedConfig;
          
          // Process data with saved hierarchy
          console.log('Processing data with saved hierarchy...');
          const processedData = processExcelData(data, savedConfig);
          
          // Render processed data
          console.log('Rendering processed data...');
          renderSheet(sheetId, processedData);
          
          // Store processed data
          window.excelData.sheetsLoaded[sheetId] = processedData;
          
          // Update export column dropdown
          populateExportColumnDropdown();
        } else {
          // No saved configuration, show modal
          console.log('No saved configuration found, showing modal...');
          showHierarchyConfigModal(data.headers, (hierarchy) => {
            console.log('Hierarchy configuration received:', hierarchy);
            
            // Store hierarchy configuration
            window.excelData.hierarchyConfigurations[sheetId] = hierarchy;
            
            // Save to localStorage
            saveHierarchyConfiguration(configKey, hierarchy);
            
            // Process data with hierarchy
            console.log('Processing data with hierarchy...');
            const processedData = processExcelData(data, hierarchy);
            
            // Render processed data
            console.log('Rendering processed data...');
            renderSheet(sheetId, processedData);
            
            // Store processed data
            window.excelData.sheetsLoaded[sheetId] = processedData;
            
            // Update export column dropdown
            populateExportColumnDropdown();
          });
        }
      } else {
        // Already processed data
        console.log('Rendering pre-processed data...');
        renderSheet(sheetId, data);
        window.excelData.sheetsLoaded[sheetId] = data;
        
        // Update export column dropdown
        populateExportColumnDropdown();
      }
    } catch (error) {
      console.error('Error loading sheet data:', error);
      const sheetContent = document.getElementById(`sheet-${sheetId}`);
      sheetContent.innerHTML = `<div class="error-message"><p>Error loading data: ${error.message}</p></div>`;
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
  
  // Render hierarchical data to the sheet
  function renderSheet(sheetId, data) {
    const sheetContent = document.getElementById(`sheet-${sheetId}`);
    
    // Create hierarchical container
    const container = document.createElement('div');
    container.className = 'excel-hierarchical-container';
    
    // Add controls 
    const controls = document.createElement('div');
    controls.className = 'controls';
    
    const expandAllButton = document.createElement('button');
    expandAllButton.textContent = 'Expand All';
    expandAllButton.className = 'control-button expand-all';
    expandAllButton.addEventListener('click', () => {
      const nodes = container.querySelectorAll('.excel-node');
      nodes.forEach(node => node.classList.remove('collapsed'));
    });
    
    const collapseAllButton = document.createElement('button');
    collapseAllButton.textContent = 'Collapse All';
    collapseAllButton.className = 'control-button collapse-all';
    collapseAllButton.addEventListener('click', () => {
      const nodes = container.querySelectorAll('.excel-node');
      nodes.forEach(node => {
        if (node.classList.contains('level-0') || node.classList.contains('level-1')) {
          node.classList.add('collapsed');
        }
      });
    });
    
    controls.appendChild(expandAllButton);
    controls.appendChild(collapseAllButton);
    container.appendChild(controls);
    
    // Add hierarchical content
    const content = document.createElement('div');
    content.className = 'excel-hierarchical-content';
    
    // Handle empty data
    if (!data.root || !data.root.children || data.root.children.length === 0) {
      content.innerHTML = '<div class="empty-sheet"><p>No data to display</p></div>';
      container.appendChild(content);
      sheetContent.innerHTML = '';
      sheetContent.appendChild(container);
      return;
    }
    
    // Render children
    data.root.children.forEach(child => {
      content.appendChild(renderNode(child, 0));
    });
    
    container.appendChild(content);
    sheetContent.innerHTML = '';
    sheetContent.appendChild(container);
  }
  
  // Render a node and its children
  function renderNode(node, level) {
    const nodeEl = document.createElement('div');
    
    // Use level-specific class names
    nodeEl.className = `level-${level}-node`;
    if (node.isEmpty) {
      nodeEl.classList.add(`level-${level}-empty`);
    }
    
    // Add data attributes for styling and debugging
    nodeEl.setAttribute('data-level', level);
    nodeEl.setAttribute('data-column-name', node.columnName || '');
    nodeEl.setAttribute('data-column-index', node.columnIndex || '');
    nodeEl.setAttribute('data-is-empty', node.isEmpty ? 'true' : 'false');
    
    if (level < 2) {
      nodeEl.classList.add(`level-${level}-collapsed`);
    }
    
    // Node header - contains column title AND cell data if this is a parent node
    const header = document.createElement('div');
    header.className = `level-${level}-header`;
    if (node.isEmpty) {
      header.classList.add(`level-${level}-empty-header`);
    }
    
    // Add expand button for level 0 and 1
    if (level < 2) {
      const expandBtn = document.createElement('div');
      expandBtn.className = `level-${level}-expand-button`;
      header.appendChild(expandBtn);
    }
    
    // Title shows the column name
    const title = document.createElement('div');
    title.className = `level-${level}-title`;
    
    // IMPORTANT: For parent nodes, show both column name AND value in the header
    // Check if node has children to determine if it's a parent
    const isParent = node.children && node.children.length > 0;
    
    if (isParent) {
      // For parent nodes: title shows both column name and cell value
      title.innerHTML = `<span class="column-name">${node.columnName}:</span> <span class="cell-value">${node.value || ''}</span>`;
    } else {
      // For leaf nodes: title shows just column name
      title.textContent = node.columnName;
    }
    
    header.appendChild(title);
    
    // Add click handler for expanding/collapsing
    header.addEventListener('click', (e) => {
      if (level < 2) {
        nodeEl.classList.toggle(`level-${level}-collapsed`);
        e.stopPropagation();
      }
    });
    
    nodeEl.appendChild(header);
    
    // Content section
    const content = document.createElement('div');
    content.className = `level-${level}-content`;
    
    if (isParent) {
      // For parent nodes: content contains children
      if (node.children && node.children.length > 0) {
        const childrenContainer = document.createElement('div');
        childrenContainer.className = `level-${level}-children`;
        childrenContainer.setAttribute('data-child-count', node.children.length);
        
        // Add class based on number of children to allow CSS targeting
        childrenContainer.classList.add(`level-${level}-child-count-${node.children.length}`);
        
        // Add layout class specific to level
        if (level === 0) {
          childrenContainer.classList.add('level-0-vertical-layout');
        } else if (level === 1) {
          childrenContainer.classList.add('level-1-grid-layout');
          
          // If there are many children, optimize the layout
          if (node.children.length > 3) {
            childrenContainer.classList.add('level-1-many-children');
          }
        }
        
        node.children.forEach(child => {
          childrenContainer.appendChild(renderNode(child, level + 1));
        });
        
        content.appendChild(childrenContainer);
      }
    } else {
      // For leaf nodes: content contains cell value
      content.textContent = node.value || '';
      
      if (node.isEmpty) {
        content.classList.add(`level-${level}-empty-content`);
      }
    }
    
    nodeEl.appendChild(content);
    
    // Process properties (non-hierarchy columns)
    if (node.properties && node.properties.length > 0) {
      const propsContainer = document.createElement('div');
      propsContainer.className = `level-${level}-properties`;
      
      // Create a row of property nodes
      node.properties.forEach(prop => {
        // Skip "Column 9" and "Column 10" properties
        if (prop.columnName === "Column 9" || prop.columnName === "Column 10") return;
        
        // Create a property node
        const propLevel = level + 1;
        const propNode = document.createElement('div');
        propNode.className = `level-${propLevel}-node`;
        propNode.setAttribute('data-column-name', prop.columnName);
        propNode.setAttribute('data-column-index', prop.columnIndex);
        
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
    
    // Add event listeners for collapse/expand all buttons
    if (collapseAllButton) {
      collapseAllButton.addEventListener('click', () => toggleAllNodes(true));
    }
    if (expandAllButton) {
      expandAllButton.addEventListener('click', () => toggleAllNodes(false));
    }
    
    // Apply initial styles
    applyStyles();
  }
  
  initialize();
}); 