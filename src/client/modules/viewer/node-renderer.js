/**
 * Node Renderer Module - HTML Node Rendering & Child Element Modals
 * Verantwoordelijk voor: Node HTML rendering, properties display, add child modals
 */

// Main function to render a node as HTML element
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
  nodeEl.setAttribute('data-is-template', node.isTemplate ? 'true' : 'false');
  nodeEl.setAttribute('data-is-duplicate', node.isDuplicate ? 'true' : 'false');
  
  // Set initial collapse state based on level and expandStateMap
  let shouldBeExpanded = expandStateMap[node.id] !== undefined ? 
    expandStateMap[node.id] : (nodeLevel >= 2);
  if (node.forceExpanded) {
    shouldBeExpanded = true;
    // Reset na render zodat het eenmalig is
    node.forceExpanded = false;
  }
  
  if (!shouldBeExpanded) {
    nodeEl.classList.add(`level-${cssLevel}-collapsed`);
  }
  
  // Node header - contains column title AND cell data if this is a parent node
  const header = document.createElement('div');
  header.className = `level-${cssLevel}-header`;
  if (node.isEmpty) {
    header.classList.add(`level-${cssLevel}-empty-header`);
  }
  
  // Add expand button for ALL parent nodes (nodes with children)
  const hasChildren = node.children && node.children.length > 0;
  if (hasChildren) {
    const expandBtn = document.createElement('div');
    expandBtn.className = `level-${cssLevel}-expand-button`;
    
    // Add click handler directly to the expand button
    expandBtn.addEventListener('click', (e) => {
      nodeEl.classList.toggle(`level-${cssLevel}-collapsed`);
      e.stopPropagation();
    });
    
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
  
  // 🔧 FIX: Title should ALWAYS be the column name (if available)
  const columnName = node.columnName || '';
  const templateIndicator = node.isTemplate ? ' 🏗️' : '';
  
  // Toon altijd de waarde uit de data als hoofdlabel
  title.textContent = (node.value || '') + templateIndicator;

  // (optioneel) Tooltip met kolomnaam
  if (columnName && columnName.trim() !== '') {
    title.title = columnName;
  }
  title.style.display = 'block';
  
  header.appendChild(title);
  
  // Use the unified edit controls from live-editor.js
  if (window.addEditControls) {
    window.addEditControls(node, nodeEl, header);
  } else {
    // Fallback if live-editor.js is not loaded
    console.warn('addEditControls not available, using fallback');
  }
  
  // Manage Items button is now handled by addEditControls function
  
  // Add click handler for expanding/collapsing - only if clicked on header itself, not on controls
  header.addEventListener('click', (e) => {
    // Only handle click if it's on the header itself or title, not on edit controls or expand button
    if ((e.target === header || e.target.classList.contains(`level-${cssLevel}-title`)) && 
        (node.children && node.children.length > 0)) {
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
    
    // Fallback to regular getExcelCoordinates (should be available from other modules)
    console.log("Falling back to getExcelCoordinates function");
    if (window.ExcelViewerNodeManager && window.ExcelViewerNodeManager.getExcelCoordinates) {
      const coordinates = window.ExcelViewerNodeManager.getExcelCoordinates(parentNode, selectedColIndex);
      
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
    } else {
      previewContent.innerHTML = '<p class="error">getExcelCoordinates function not available</p>';
    }
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
    if (!coordinates && window.ExcelViewerNodeManager && window.ExcelViewerNodeManager.getExcelCoordinates) {
      coordinates = window.ExcelViewerNodeManager.getExcelCoordinates(parentNode, selectedColIndex);
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
    
    // Add node and re-render with preserved expand state (should be available from node-management module)
    if (window.ExcelViewerNodeManager && window.ExcelViewerNodeManager.addNode) {
      window.ExcelViewerNodeManager.addNode(parentNode, newNode);
    } else {
      console.error("addNode function not available");
    }
  });
  
  // Handle cancel
  cancelButton.addEventListener('click', () => {
    document.body.removeChild(modal);
  });
}

// Export functions for other modules
window.ExcelViewerNodeRenderer = {
  renderNode,
  showAddChildElementModal
}; 