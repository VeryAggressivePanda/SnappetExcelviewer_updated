/**
 * Data Processing Module - Excel Data Processing & Hierarchy Building
 * Verantwoordelijk voor: Hierarchy config modal, Excel data processing, node creation
 */

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
      
      // 🔧 FIX: Validate childColIndex before creating node
      if (childColIndex === undefined || childColIndex === null || isNaN(childColIndex)) {
        console.error('❌ Invalid childColIndex in data-processing:', childColIndex);
        continue; // Skip this invalid column
      }
      
      if (childColIndex < 0 || childColIndex >= headers.length) {
        console.error('❌ childColIndex out of range in data-processing:', childColIndex, 'Headers length:', headers.length);
        continue; // Skip this out-of-range column
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
        parent: parentNode, // ADD PARENT REFERENCE
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

// Export functions for other modules
window.ExcelViewerDataProcessor = {
  showHierarchyConfigModal,
  processExcelData,
  processChildrenWithCurrentValues,
  cleanupNode,
  addPropertiesToNode
}; 