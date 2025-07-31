/**
 * Node Management Module - Node CRUD Operations & Excel Coordinates
 * Verantwoordelijk voor: Excel coordinates, node add/delete, expand state, hierarchy utilities
 */

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

  // --- FORCEER EXPANDED STATE OP PARENT ---
  activeNode.forceExpanded = true;

  // --- AUTO-EXPAND PARENT ---
  const expandStateMap = {};
  if (activeNode.id) {
    expandStateMap[activeNode.id] = true;
  }

  // Re-render sheet (should be available from other modules)
  if (window.ExcelViewerSheetManager && window.ExcelViewerSheetManager.renderSheet) {
    window.ExcelViewerSheetManager.renderSheet(activeSheetId, sheetData, expandStateMap);
  }
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

// Function to add a new child node and preserve expand/collapse state (IMPROVED TEMPLATE AWARE)
function addNode(parentNode, newNode) {
  // Check if this is a duplicate - only templates can be modified
  if (parentNode.isDuplicate) {
    alert('🚫 This is a duplicate node. Please modify the template (- master) instead.');
    return;
  }
  
  if (!parentNode.children) {
    parentNode.children = [];
  }

  // Store expand state before modifying the DOM
  const expandStateMap = storeExpandState();

  // ADD PARENT REFERENCE to new node
  newNode.parent = parentNode;

  // Add the new node to the parent
  parentNode.children.push(newNode);

  // --- FORCEER EXPANDED STATE OP PARENT ---
  parentNode.forceExpanded = true;

  // --- AUTO-EXPAND PARENT ---
  if (parentNode.id) {
    expandStateMap[parentNode.id] = true;
  }

  // Mark parent as GLOBAL template if it has nodes of same type anywhere (use live-editor function)
  if (window.markGlobalTemplatesOfType) {
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    const rootNode = sheetData.root;
    window.markGlobalTemplatesOfType(rootNode, parentNode.columnName, parentNode.columnIndex, parentNode.id);
  }

  // Global template replication: copy structure to ALL nodes of same type
  if (parentNode.isTemplate && window.applyGlobalTemplateReplication) {
    window.applyGlobalTemplateReplication(parentNode);
  } else {
    console.log('🔧 SMART: Template replication with contextual Excel data');
    // Re-render with preserved expand state, but force parent expanded
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    if (window.ExcelViewerSheetManager && window.ExcelViewerSheetManager.renderSheet) {
      window.ExcelViewerSheetManager.renderSheet(activeSheetId, sheetData, expandStateMap);
    }
  }
}

// Function to delete a node (IMPROVED TEMPLATE AWARE)
function deleteNode(parentNode, nodeToDelete) {
  // Check if this is a duplicate - only templates can be modified
  if (nodeToDelete.isDuplicate) {
    alert('🚫 This is a duplicate node. Please modify the template (- master) instead.');
    return;
  }
  
  // Check if this is a template - warn about deleting template
  if (nodeToDelete.isTemplate) {
    if (!confirm(`⚠️ This is a TEMPLATE node (- master). Deleting it will affect ALL duplicate nodes.\n\nDelete template "${nodeToDelete.value || nodeToDelete.columnName || 'New Container'}" and all its children?`)) {
      return;
    }
  } else {
    if (!confirm(`Delete node "${nodeToDelete.value || nodeToDelete.columnName || 'New Container'}" and all its children?`)) {
      return;
    }
  }
  
  if (!parentNode.children) return;
  
  // Store expand state before modifying the DOM
  storeExpandState();
  
  // Remove the node from the parent
  parentNode.children = parentNode.children.filter(child => child.id !== nodeToDelete.id);
  
  // Smart template replication: copy structure but use contextual Excel data
  if (parentNode.isTemplate && window.applySmartTemplateReplication) {
    window.applySmartTemplateReplication(parentNode);
  } else {
    // Re-render the node
    const activeSheetId = window.excelData.activeSheetId;
    const sheetContainer = document.getElementById('sheet-container');
    if (sheetContainer && window.ExcelViewerNodeRenderer && window.ExcelViewerNodeRenderer.renderNode) {
      sheetContainer.innerHTML = '';
      window.ExcelViewerNodeRenderer.renderNode(window.hierarchicalData, 0);
      if (window.ExcelViewerUIControls && window.ExcelViewerUIControls.applyStyles) {
        window.ExcelViewerUIControls.applyStyles();
      }
    }
  }
}

// Helper function to ensure parent references are maintained throughout the tree
function updateParentReferences(node, parent = null) {
  if (!node) return;
  
  // Set parent reference
  node.parent = parent;
  
  // Recursively update children
  if (node.children && node.children.length > 0) {
    node.children.forEach(child => {
      updateParentReferences(child, node);
    });
  }
}

// GLOBAL: Helper function to find and mark templates across entire tree by type
function identifyAndMarkGlobalTemplates(rootNode) {
  if (!rootNode) return;
  
  // Collect ALL nodes across entire tree grouped by type (columnName + columnIndex)
  const globalTypeGroups = {};
  
  function collectAllNodes(node) {
    if (node.columnName && node.columnIndex !== undefined) {
      const key = `${node.columnName}-${node.columnIndex}`;
      if (!globalTypeGroups[key]) {
        globalTypeGroups[key] = [];
      }
      globalTypeGroups[key].push(node);
    }
    if (node.children && node.children.length > 0) {
      node.children.forEach(collectAllNodes);
    }
  }
  
  collectAllNodes(rootNode);
  
  // Mark templates and duplicates globally
  Object.entries(globalTypeGroups).forEach(([key, nodes]) => {
    if (nodes.length > 1) {
      // First node becomes GLOBAL template
      const template = nodes[0];
      template.isTemplate = true;
      template.isDuplicate = false;
      console.log(`- master GLOBAL: Marked "${template.value}" (${template.columnName}) as GLOBAL template`);
      
      // Rest become GLOBAL duplicates
      nodes.slice(1).forEach(duplicate => {
        duplicate.isDuplicate = true;
        duplicate.isTemplate = false;
        duplicate.templateId = template.id;
        console.log(`📋 GLOBAL: Marked "${duplicate.value}" (${duplicate.columnName}) as GLOBAL duplicate`);
      });
    }
  });
}

// Export functions for other modules  
window.ExcelViewerNodeManager = {
  getExcelCoordinates,
  getExcelCellValue,
  addElementDirectly,
  getSavedHierarchyConfiguration,
  findParentNode,
  storeExpandState,
  addNode,
  deleteNode,
  updateParentReferences,
  identifyAndMarkGlobalTemplates
}; 