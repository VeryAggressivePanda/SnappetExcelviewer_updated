// Live editing functionality for Excel Viewer

// Function to enable live editing on a sheet
function enableLiveEditing(sheetId) {
  const sheetData = window.excelData.sheetsLoaded[sheetId];
  
  // If no data, start with empty structure
  if (!sheetData || !sheetData.root || sheetData.root.children.length === 0) {
    sheetData.root = {
      type: 'root',
      children: [{
        id: `node-${Date.now()}`,
        value: 'Start Building',
        columnName: '',
        level: 0,
        children: [],
        properties: [],
        isPlaceholder: true
      }],
      level: -1
    };
  }
  
  // Mark as editable
  sheetData.isEditable = true;
}

// Function to add edit controls to a node
function addEditControls(node, nodeEl, header) {
  const editControls = document.createElement('div');
  editControls.className = 'edit-controls';
  editControls.style.display = 'flex';
  editControls.style.gap = '5px';
  editControls.style.marginLeft = 'auto';
  editControls.style.position = 'relative';
  
  // Add direct "+Add Content" button for nodes without meaningful data
  // Show for: empty containers, placeholders, or nodes without column data
  const hasNoRealData = (!node.columnIndex && node.columnIndex !== 0) || 
                        node.isPlaceholder || 
                        (node.value && (node.value === 'Container' || node.value === 'New Container'));
  
  if (hasNoRealData) {
    const addContentBtn = document.createElement('button');
    addContentBtn.className = 'add-content-button';
    addContentBtn.innerHTML = '➕ Add Content';
    addContentBtn.style.padding = '6px 12px';
    addContentBtn.style.fontSize = '12px';
    addContentBtn.style.border = '2px solid #4caf50';
    addContentBtn.style.borderRadius = '1rem';
    addContentBtn.style.background = '#4caf50';
    addContentBtn.style.color = 'white';
    addContentBtn.style.cursor = 'pointer';
    addContentBtn.style.fontWeight = '500';
    addContentBtn.style.transition = 'all 0.2s ease';
    addContentBtn.style.fontFamily = 'Inter, sans-serif';
    
    addContentBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      console.log('Add Content clicked for node:', node);
      showColumnSelectionArea(node);
    });
    
    addContentBtn.addEventListener('mouseenter', () => {
      addContentBtn.style.background = '#45a049';
      addContentBtn.style.transform = 'translateY(-1px)';
      addContentBtn.style.boxShadow = '0 2px 8px rgba(76, 175, 80, 0.3)';
    });
    
    addContentBtn.addEventListener('mouseleave', () => {
      addContentBtn.style.background = '#4caf50';
      addContentBtn.style.transform = 'translateY(0)';
      addContentBtn.style.boxShadow = 'none';
    });
    
    editControls.appendChild(addContentBtn);
  }
  
  // Only create kebab menu for nodes that already have data/content
  // EXCLUDE nodes that show "+Add Content" button to avoid double controls
  if (((node.columnIndex || node.columnIndex === 0) || node.isPlaceholder) && !hasNoRealData) {
    // Menu items array - build first to check if we need menu
    const menuItems = [];
  
    // Add Sibling option (not for root)
    if (node.level > -1 && !node.isPlaceholder) {
      menuItems.push({
        text: '➕ Add Sibling',
        icon: '➕',
        action: () => addSiblingContainer(node),
        color: '#2196f3'
      });
    }
    
    // Layout toggle option (available for ALL nodes, not just those with children)
    if (!node.isPlaceholder) {
      const currentLayout = node.layoutMode || 'horizontal';
      const newLayout = currentLayout === 'horizontal' ? 'vertical' : 'horizontal';
      const layoutIcon = currentLayout === 'horizontal' ? '⬇️' : '➡️';
      
      menuItems.push({
        text: `${layoutIcon} Layout Toggle`,
        icon: layoutIcon,
        action: () => toggleNodeLayout(node),
        color: '#ff9800'
      });
    }
    
    // Manage Items option for nodes with column data
    if ((node.columnIndex || node.columnIndex === 0) && !node.isPlaceholder) {
      menuItems.push({
        text: '☑️ Refresh Values',
        icon: '🔄',
        action: () => {
          if (node.columnIndex !== undefined) {
            console.log('Refreshing column data for:', node.columnName);
            assignColumnToNode(node, node.columnIndex);
          }
        },
        color: '#9c27b0'
      });
    }
    
    // Delete option (not for root)
    if (node.level > -1 && !node.isPlaceholder) {
      menuItems.push({
        text: '🗑️ Delete',
        icon: '❌',
        action: () => deleteContainer(node),
        color: '#f44336',
        separator: true
      });
    }
    
    // Only create kebab menu if there are actually menu items to show
    if (menuItems.length > 0) {
      // Create kebab menu button
      const kebabBtn = document.createElement('button');
      kebabBtn.className = 'kebab-menu-button';
      kebabBtn.innerHTML = '⋮';
      
      // Create dropdown menu
      const dropdownMenu = document.createElement('div');
      dropdownMenu.className = 'kebab-dropdown-menu';
      dropdownMenu.style.display = 'none';
      
      // Create menu item elements
      menuItems.forEach((item, index) => {
        if (item.separator && index > 0) {
          const separator = document.createElement('div');
          separator.style.height = '1px';
          separator.style.background = '#eee';
          separator.style.margin = '4px 0';
          dropdownMenu.appendChild(separator);
        }
      
      const menuItem = document.createElement('div');
      menuItem.className = 'kebab-menu-item';
      menuItem.textContent = item.text;
      menuItem.style.padding = '8px 12px';
      menuItem.style.cursor = 'pointer';
      menuItem.style.fontSize = '13px';
      menuItem.style.color = item.color;
      menuItem.style.borderBottom = index < menuItems.length - 1 ? '1px solid #f0f0f0' : 'none';
      menuItem.style.transition = 'background-color 0.2s ease';
      
      menuItem.addEventListener('mouseenter', () => {
        menuItem.style.backgroundColor = '#f5f5f5';
      });
      
      menuItem.addEventListener('mouseleave', () => {
        menuItem.style.backgroundColor = 'transparent';
      });
      
      menuItem.addEventListener('click', (e) => {
        e.stopPropagation();
        dropdownMenu.style.display = 'none';
        item.action();
      });
      
      dropdownMenu.appendChild(menuItem);
    });
    
      // Toggle dropdown on kebab button click
      kebabBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        const isVisible = dropdownMenu.style.display === 'block';
        
        // Hide all other open dropdowns
        document.querySelectorAll('.kebab-dropdown-menu').forEach(menu => {
          menu.style.display = 'none';
        });
        
        // Toggle this dropdown
        dropdownMenu.style.display = isVisible ? 'none' : 'block';
      });
      
      // Close dropdown when clicking outside
      document.addEventListener('click', (e) => {
        if (!editControls.contains(e.target)) {
          dropdownMenu.style.display = 'none';
        }
      });
      
      // Hover effects for kebab button
      kebabBtn.addEventListener('mouseenter', () => {
        kebabBtn.style.background = '#e9ecef';
        kebabBtn.style.borderColor = '#adb5bd';
      });
      
      kebabBtn.addEventListener('mouseleave', () => {
        kebabBtn.style.background = '#f8f9fa';
        kebabBtn.style.borderColor = '#ddd';
      });
  
      editControls.appendChild(kebabBtn);
      editControls.appendChild(dropdownMenu);
    } // End of menuItems.length > 0 check
  } // End of kebab menu conditional
  
  header.appendChild(editControls);
}

// Function to add a child container (IMPROVED WITH TEMPLATE MARKING)
function addChildContainer(parentNode, skipReRender = false) {
  // Check if this is a duplicate - only templates can be modified
  if (parentNode.isDuplicate) {
    alert('🚫 This is a duplicate node. Please modify the template (- master) instead.');
    return;
  }
  
  // Remove placeholder if it exists
  if (parentNode.isPlaceholder) {
    parentNode.isPlaceholder = false;
    parentNode.value = 'Container';
  }
  
  const newNode = {
    id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
    value: 'New Container',
    columnName: '',
    level: parentNode.level + 1,
    children: [],
    properties: [],
    layoutMode: 'horizontal',
    parent: parentNode // ADD PARENT REFERENCE
  };
  
  if (!parentNode.children) {
    parentNode.children = [];
  }
  
  parentNode.children.push(newNode);

  // FORCEER EXPANDED STATE OP PARENT DIRECT NA ADD CHILD
  const parentNodeEl = document.querySelector(`[data-node-id='${parentNode.id}']`);
  if (parentNodeEl) {
    const level = parentNodeEl.getAttribute('data-level') || parentNode.level || 0;
    parentNodeEl.classList.remove(`level-${level}-collapsed`);
  }
  
  // Get root to do global template marking
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const rootNode = sheetData.root;
  
  // Mark parent as GLOBAL template if it has nodes of same type anywhere in tree
  markGlobalTemplatesOfType(rootNode, parentNode.columnName, parentNode.columnIndex);
  
  // Template replication if this is a template
  if (parentNode.isTemplate) {
    applyGlobalTemplateReplication(parentNode);
  } else if (!skipReRender) {
    // Just re-render normally for now
    const activeSheetId = window.excelData.activeSheetId;
    if (window.renderSheet) {
      window.renderSheet(activeSheetId, window.excelData.sheetsLoaded[activeSheetId]);
    }
    applyFlexboxStyling();
  }
}

// Function to add a sibling container (IMPROVED WITH TEMPLATE MARKING)
function addSiblingContainer(node) {
  // Check if this is a duplicate - only templates can be modified
  if (node.isDuplicate) {
    alert('🚫 This is a duplicate node. Please modify the template (- master) instead.');
    return;
  }
  
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  // Find parent
  const parent = findParentInStructure(sheetData.root, node.id);
  
  const newNode = {
    id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
    value: 'New Container',
    columnName: '',
    level: node.level,
    children: [],
    properties: [],
    layoutMode: 'horizontal',
    parent: parent // ADD PARENT REFERENCE
  };
  
  if (!parent.children) {
    parent.children = [];
  }
  
  // Insert after current node
  const currentIndex = parent.children.findIndex(child => child.id === node.id);
  parent.children.splice(currentIndex + 1, 0, newNode);
  
  console.log('- master Sibling added');
  
  // Get root to do global template marking
  const rootNode = sheetData.root;
  
  // Mark original node as GLOBAL template if it has nodes of same type anywhere in tree
  markGlobalTemplatesOfType(rootNode, node.columnName, node.columnIndex);
  
  // Unified replication if the original node is a template
  if (node.isTemplate) {
          applyGlobalTemplateReplication(node);
  } else {
    // Re-render normally
    if (window.renderSheet) {
      window.renderSheet(activeSheetId, sheetData);
    }
    applyFlexboxStyling();
  }
}

// Function to delete a container
function deleteContainer(node) {
  // Check if this is a duplicate - only templates can be modified
  if (node.isDuplicate) {
    alert('🚫 This is a duplicate node. Please modify the template (- master) instead.');
    return;
  }
  
  // Check if this is a template - warn about deleting template
  if (node.isTemplate) {
    if (!confirm(`⚠️ This is a TEMPLATE node (- master). Deleting it will affect ALL duplicate nodes.\n\nDelete template "${node.value || node.columnName || 'New Container'}" and all its children?`)) {
      return;
    }
  } else {
    if (!confirm(`Delete container "${node.value || node.columnName || 'New Container'}" and all its children?`)) {
      return;
    }
  }
  
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  // Find parent and remove node
  const parent = findParentInStructure(sheetData.root, node.id);
  if (parent && parent.children) {
    parent.children = parent.children.filter(child => child.id !== node.id);
  }
  
  // Re-render
  if (window.renderSheet) {
    window.renderSheet(activeSheetId, sheetData);
  }
  applyFlexboxStyling();
}

// Function to assign a column to a node (IMPROVED VERSION!)
function assignColumnToNode(node, columnIndex, skipReplication = false) {
  // Validate node first
  if (!node) {
    console.error('❌ Node is null or undefined in assignColumnToNode');
    return;
  }
  
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const headers = sheetData.headers;
  
  // Validate column index first
  if (columnIndex === undefined || columnIndex === null || isNaN(columnIndex)) {
    console.error('❌ Invalid columnIndex in assignColumnToNode:', columnIndex);
    return;
  }
  
  if (columnIndex < 0 || columnIndex >= headers.length) {
    console.error('❌ ColumnIndex out of range in assignColumnToNode:', columnIndex, 'Headers length:', headers.length);
    return;
  }
  
  // Update node with column info
  node.columnIndex = columnIndex;
  node.columnName = headers[columnIndex] || `Column ${columnIndex + 1}`;
  
  // Get contextually filtered values
  const contextualValues = getContextualValuesForNode(node, columnIndex);
  
  if (contextualValues.length === 0) {
    // No values found → LEAVE EMPTY, don't use column name!
    node.value = '';
    node.children = [];
  } else if (contextualValues.length === 1) {
    // Single value - assign it directly
    node.value = contextualValues[0];
    node.children = [];
  } else {
    // Multiple values - CREATE SIBLINGS AUTOMATICALLY
    createAutoSiblings(node, contextualValues);
    return; // Exit early, createAutoSiblings handles everything
  }
  
  // Only re-render if not during replication
  if (!skipReplication) {
    // Regular re-render - template replication is handled in createAutoSiblings
    if (window.renderSheet) {
      window.renderSheet(activeSheetId, sheetData);
    }
    applyFlexboxStyling();
  }
}

// NEW: Function to assign a single specific value to a node (not all unique values)
function assignSpecificValueToNode(node, columnIndex) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const headers = sheetData.headers;
  
  // Validate column index first
  if (columnIndex === undefined || columnIndex === null || isNaN(columnIndex)) {
    console.error('❌ Invalid columnIndex:', columnIndex);
    return;
  }
  
  if (columnIndex < 0 || columnIndex >= headers.length) {
    console.error('❌ ColumnIndex out of range:', columnIndex, 'Headers length:', headers.length);
    return;
  }
  
  // Update node with column info
  node.columnIndex = columnIndex;
  node.columnName = headers[columnIndex] || `Column ${columnIndex + 1}`;
  
  // Get the SPECIFIC row data for this node's context
  const specificRowData = getSpecificRowDataForNode(node);
  
  if (specificRowData && specificRowData[columnIndex] !== undefined && specificRowData[columnIndex] !== null && specificRowData[columnIndex] !== '') {
    // Set the specific value from that row
    node.value = specificRowData[columnIndex].toString().trim();
  } else {
    // No data found → LEAVE EMPTY, don't repeat column name!
    node.value = '';
  }
  
  // No children needed for specific values
  node.children = [];
  
  // 🚫 NO template replication here - let the parent handle it in createMultipleColumnContainers
  // This prevents double replication and conflicts
  
  // 🚫 NO re-render here - let the parent handle it to avoid multiple renders
}

// NEW: Function to get specific row data for a node's context
function getSpecificRowDataForNode(node) {
  const activeSheetId = window.excelData.activeSheetId;
  const rawData = getRawExcelData(activeSheetId);
  
  if (!rawData || rawData.length <= 1) {
    return null;
  }
  
  // Get the full parent context (Blok 1 → Week 1 → Les 1)
  const fullParentContext = getFullParentContext(node);
  
  // IMPORTANT: If this node itself is hierarchical (like Les 1), add it to the context for filtering
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const headers = sheetData.headers;
  
  if (node.columnIndex !== undefined && node.columnName && node.value) {
    // Add ALL nodes with column info to context (not just strict hierarchical patterns)
    // This includes "Les 1", "Speelles: de tijd gaat door", "4 x 10 minuten...", etc.
    const shouldAddToContext = node.columnName === 'Blok' || node.columnName === 'Week' || node.columnName === 'Les';
    
    if (shouldAddToContext) {
      fullParentContext.push({
        columnIndex: node.columnIndex,
        columnName: node.columnName,
        value: node.value
      });
    }
  }
  
  // Find the specific row that matches this exact context
  const matchingRows = findRowsForParentHierarchy(rawData, fullParentContext);
  
  if (matchingRows.length > 0) {
    // Return the first matching row
    const specificRow = matchingRows[0];
    return specificRow;
  } else {
    return null;
  }
}

// IMPROVED function to create siblings automatically with proper template marking
function createAutoSiblings(node, values) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  // Find parent and update parent references
  const parent = findParentInStructure(sheetData.root, node.id);
  if (!parent || !parent.children) return;
  
  // Find position of current node
  const currentIndex = parent.children.findIndex(child => child.id === node.id);
  if (currentIndex === -1) return;
  

  
  // Update current node with first value and mark as template
  node.value = values[0];
  node.children = [];
  node.isTemplate = true;
  node.parent = parent; // ADD PARENT REFERENCE
  
  // Create siblings for remaining values
  const newSiblings = values.slice(1).map(value => ({
    id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
    value: value,
    columnName: node.columnName,
    columnIndex: node.columnIndex,
    level: node.level,
    children: [],
    properties: [],
    layoutMode: node.layoutMode || 'horizontal',
    isDuplicate: true,
    templateId: node.id,
    parent: parent // ADD PARENT REFERENCE
  }));
  
  // Insert siblings after current node
  parent.children.splice(currentIndex + 1, 0, ...newSiblings);
  
  console.log(`✅ Created template "${values[0]}" with ${newSiblings.length} duplicates`);
  
  // Check if parent is a template and trigger replication
  if (parent && parent.isTemplate) {
    console.log('🔄 Parent is template - triggering GLOBAL replication after sibling creation');
    applyGlobalTemplateReplication(parent);
  } else {
    // Regular re-render
    if (window.renderSheet) {
      window.renderSheet(activeSheetId, sheetData);
    }
    applyFlexboxStyling();
  }
}

// Function to show choice between siblings or children for multiple values
function showSiblingOrChildrenChoice(node, selectedValues) {
  const modal = document.createElement('div');
  modal.className = 'pdf-preview-modal';
  
  const content = document.createElement('div');
  content.className = 'pdf-preview-content';
  content.style.maxWidth = '400px';
  content.style.padding = '20px';
  
  const header = document.createElement('h2');
  header.textContent = 'Multiple Values Selected';
  header.style.marginBottom = '20px';
  
  const description = document.createElement('p');
  description.textContent = `You selected ${selectedValues.length} values. How would you like to organize them?`;
  description.style.marginBottom = '20px';
  
  const optionsContainer = document.createElement('div');
  optionsContainer.style.marginBottom = '20px';
  
  // Option 1: Create as siblings
  const siblingsOption = document.createElement('div');
  siblingsOption.style.marginBottom = '15px';
  siblingsOption.style.padding = '10px';
  siblingsOption.style.border = '2px solid #e0e0e0';
  siblingsOption.style.borderRadius = '5px';
  siblingsOption.style.cursor = 'pointer';
  siblingsOption.style.transition = 'border-color 0.2s';
  
  const siblingsTitle = document.createElement('h3');
  siblingsTitle.textContent = '📄 Create as Siblings';
  siblingsTitle.style.margin = '0 0 5px 0';
  siblingsTitle.style.color = '#2196f3';
  
  const siblingsDesc = document.createElement('p');
  siblingsDesc.textContent = 'Each value becomes a separate container next to each other';
  siblingsDesc.style.margin = '0';
  siblingsDesc.style.fontSize = '14px';
  siblingsDesc.style.color = '#666';
  
  siblingsOption.appendChild(siblingsTitle);
  siblingsOption.appendChild(siblingsDesc);
  
  // Option 2: Create as children
  const childrenOption = document.createElement('div');
  childrenOption.style.padding = '10px';
  childrenOption.style.border = '2px solid #e0e0e0';
  childrenOption.style.borderRadius = '5px';
  childrenOption.style.cursor = 'pointer';
  childrenOption.style.transition = 'border-color 0.2s';
  
  const childrenTitle = document.createElement('h3');
  childrenTitle.textContent = '📁 Create as Children';
  childrenTitle.style.margin = '0 0 5px 0';
  childrenTitle.style.color = '#4caf50';
  
  const childrenDesc = document.createElement('p');
  childrenDesc.textContent = 'All values become children inside this container';
  childrenDesc.style.margin = '0';
  childrenDesc.style.fontSize = '14px';
  childrenDesc.style.color = '#666';
  
  childrenOption.appendChild(childrenTitle);
  childrenOption.appendChild(childrenDesc);
  
  optionsContainer.appendChild(siblingsOption);
  optionsContainer.appendChild(childrenOption);
  
  const buttonContainer = document.createElement('div');
  buttonContainer.style.display = 'flex';
  buttonContainer.style.justifyContent = 'center';
  buttonContainer.style.marginTop = '20px';
  
  const cancelButton = document.createElement('button');
  cancelButton.textContent = 'Cancel';
  cancelButton.className = 'control-button';
  
  buttonContainer.appendChild(cancelButton);
  
  content.appendChild(header);
  content.appendChild(description);
  content.appendChild(optionsContainer);
  content.appendChild(buttonContainer);
  modal.appendChild(content);
  document.body.appendChild(modal);
  
  // Add hover effects
  siblingsOption.addEventListener('mouseenter', () => {
    siblingsOption.style.borderColor = '#2196f3';
  });
  siblingsOption.addEventListener('mouseleave', () => {
    siblingsOption.style.borderColor = '#e0e0e0';
  });
  
  childrenOption.addEventListener('mouseenter', () => {
    childrenOption.style.borderColor = '#4caf50';
  });
  childrenOption.addEventListener('mouseleave', () => {
    childrenOption.style.borderColor = '#e0e0e0';
  });
  
  // Event listeners
  cancelButton.addEventListener('click', () => {
    document.body.removeChild(modal);
  });
  
  siblingsOption.addEventListener('click', () => {
    document.body.removeChild(modal);
    createAsSiblings(node, selectedValues);
  });
  
  childrenOption.addEventListener('click', () => {
    document.body.removeChild(modal);
    createAsChildren(node, selectedValues);
  });
}

// Function to create template system with siblings (TEMPLATE SYSTEM)
// Removed - using simple template system now

// Removed complex template functions - using simple approach now

// Removed complex data population - using simple approach now

// Removed all complex legacy functions - keeping it simple now

// Function to find a node by ID in the data structure
function findNodeById(nodeId) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  if (!sheetData || !sheetData.root) return null;
  
  function searchNode(node) {
    if (node.id === nodeId) return node;
    
    if (node.children) {
      for (const child of node.children) {
        const found = searchNode(child);
        if (found) return found;
      }
    }
    
    return null;
  }
  
  return searchNode(sheetData.root);
}

// Function to apply proper flexbox styling after rendering (simplified)
function applyFlexboxStyling() {
  // Use the new layout styling function instead
  applyLayoutStyling();
}

// Function to get used columns in the hierarchy path to this node
function getUsedColumnsInHierarchy(node) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const usedColumns = [];
  
  // 1. Traverse up the hierarchy to find all used column indices in parent chain
  let currentNode = node;
  while (currentNode && currentNode.level >= 0) {
    const parent = findParentInStructure(sheetData.root, currentNode.id);
    if (parent && (parent.columnIndex || parent.columnIndex === 0)) {
      usedColumns.push(parent.columnIndex);
    }
    currentNode = parent;
  }
  
  // 2. Check sibling nodes on the same level (template system consideration)
  const parent = findParentInStructure(sheetData.root, node.id);
  if (parent && parent.children) {
    parent.children.forEach(sibling => {
      if (sibling.id !== node.id && (sibling.columnIndex || sibling.columnIndex === 0)) {
        usedColumns.push(sibling.columnIndex);
      }
    });
  }
  
  // 3. Check children of current node (already used columns by children)
  function collectChildrenColumns(nodeToCheck) {
    if (nodeToCheck.children) {
      nodeToCheck.children.forEach(child => {
        if (child.columnIndex || child.columnIndex === 0) {
          usedColumns.push(child.columnIndex);
        }
        // Recursively check grandchildren
        collectChildrenColumns(child);
      });
    }
  }
  collectChildrenColumns(node);
  
  // 4. If this node is part of template system, check template structure
  // If the current node is a duplicate, check what the template uses
  if (node.isDuplicate && node.templateId) {
    const templateNode = findNodeById(node.templateId);
    if (templateNode) {
      collectChildrenColumns(templateNode);
    }
  }
  
  // Remove duplicates and return
  return [...new Set(usedColumns)];
}

// REMOVED: Duplicate function - using improved version above

// IMPROVED: Get contextual values that properly filters by parent hierarchy
function getContextualValuesForNode(node, columnIndex) {
  const activeSheetId = window.excelData.activeSheetId;
  const rawData = getRawExcelData(activeSheetId);
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const headers = sheetData.headers;
  const columnName = headers[columnIndex] || `Column ${columnIndex + 1}`;
  
  if (!rawData || rawData.length <= 1) {
    return [];
  }
  
  // Get the FULL parent context (all parents up the hierarchy)
  const fullParentContext = getFullParentContext(node);
  
  // Find rows that match this parent hierarchy (if any)
  const matchingRows = findRowsForParentHierarchy(rawData, fullParentContext);
  
  // Get ONLY values that actually exist for THIS specific parent context
  const uniqueValues = new Set();
  
  matchingRows.forEach((row, index) => {
    const cellValue = row[columnIndex];
    
    if (cellValue && cellValue.toString().trim()) {
      const value = cellValue.toString().trim();
      uniqueValues.add(value);
    }
  });
  
  let result = Array.from(uniqueValues).sort();
  
  // HIERARCHICAL COLUMN FIX: Only filter Blok and Week, keep ALL Les values
  const isStrictHierarchicalColumn = columnName === 'Blok' || columnName === 'Week';
  if (isStrictHierarchicalColumn) {
    // For Blok and Week columns, filter strictly for hierarchy labels
    const hierarchicalValues = result.filter(value => {
      if (columnName === 'Blok') {
        return value.match(/^Blok\s+\d+$/i);
      } else if (columnName === 'Week') {
        return value.match(/^week\s+\d+$/i);
      }
      return false;
    });
    
    if (hierarchicalValues.length > 0) {
      result = hierarchicalValues;
    }
  }
  
  return result;
}

// NEW: Function specifically for duplicate node context filtering
function getContextualValuesForDuplicateNode(parentContext, filterValue, filterColumnIndex, targetColumnIndex) {
  const activeSheetId = window.excelData.activeSheetId;
  const rawData = getRawExcelData(activeSheetId);
  
  if (!rawData || rawData.length <= 1) {
    return [];
  }
  
  // Step 1: Find row range for parent context
  let rowRange = [];
  if (parentContext.length === 0) {
    // No parent context - use all rows
    rowRange = rawData.slice(1); // Skip header
  } else {
    // Find rows matching parent hierarchy
    rowRange = findRowsForParentHierarchy(rawData, parentContext);
  }
  
  // Step 2: Within that range, filter for the specific filter value (handle merged cells!)
  // MERGED CELL FIX: Find all rows that belong to this filter value
  const matchingRows = [];
  let currentFilterValue = null;
  let isInTargetGroup = false;
  
  for (let i = 0; i < rowRange.length; i++) {
    const row = rowRange[i];
    const cellValue = row[filterColumnIndex];
    
    if (cellValue && cellValue.toString().trim()) {
      // Non-empty cell - this defines the current group
      currentFilterValue = cellValue.toString().trim();
      isInTargetGroup = (currentFilterValue === filterValue);
      
      if (isInTargetGroup) {
        matchingRows.push(row);
      }
    } else if (isInTargetGroup) {
      // Empty cell, but we're in the target group (merged cell)
      matchingRows.push(row);
    }
  }
  
  if (matchingRows.length === 0) {
    return [];
  }
  
  // Step 3: Extract unique values from the target column within filtered rows
  const uniqueValues = new Set();
  matchingRows.forEach((row, index) => {
    const value = row[targetColumnIndex];
    if (value && value.toString().trim()) {
      uniqueValues.add(value.toString().trim());
    }
  });
  
  const result = Array.from(uniqueValues).sort();
  
  return result;
}

// IMPROVED: Get full parent context with better hierarchy tracking
function getFullParentContext(node) {
  const context = [];
  let currentNode = node.parent; // Start from parent, not the node itself
  
  // Traverse up the hierarchy
  while (currentNode && currentNode.columnIndex !== undefined) {
    const parentInfo = {
      columnIndex: currentNode.columnIndex,
      columnName: currentNode.columnName,
      value: currentNode.value
    };
    context.unshift(parentInfo);
    currentNode = currentNode.parent;
  }
  
  return context;
}

// IMPROVED: Find rows that match the EXACT parent hierarchy  
function findRowsForParentHierarchy(rawData, parentContext) {
  if (!rawData || rawData.length <= 1 || parentContext.length === 0) {
    return rawData.slice(1);
  }
  
  const dataRows = rawData.slice(1); // Skip header
  const matchingRows = [];
  
  // Track running values for merged cells (Excel pattern)
  const runningValues = {};
  
  // Find ALL rows that match the parent hierarchy, handling merged cells
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    
    // Update running values from this row (handle merged cells)
    for (let colIdx = 0; colIdx < Math.max(row.length, 10); colIdx++) {
      if (row[colIdx] && row[colIdx].toString().trim()) {
        runningValues[colIdx] = row[colIdx].toString().trim();
      }
    }
    
    // Check if current running values match ALL parent context requirements
    let matchesAll = true;
    for (const parentCtx of parentContext) {
      const currentValue = runningValues[parentCtx.columnIndex];
      const matches = currentValue && currentValue === parentCtx.value;
      if (!matches) {
        matchesAll = false;
        break;
      }
    }
    
    if (matchesAll) {
      matchingRows.push(row);
    }
  }
  
  return matchingRows;
}

// Function to show inline value selector with checkboxes (NO MODAL!)
function showInlineValueSelector(node, columnIndex, editControls, oldSelect) {
  console.log('showInlineValueSelector called with:', { node, columnIndex, editControls, oldSelect });
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const headers = sheetData.headers;
  
  // 🔧 FIX: Validate column index before assignment
  if (columnIndex === undefined || columnIndex === null || isNaN(columnIndex)) {
    console.error('❌ Invalid columnIndex in showInlineValueSelector:', columnIndex);
    return;
  }
  
  if (columnIndex < 0 || columnIndex >= headers.length) {
    console.error('❌ ColumnIndex out of range in showInlineValueSelector:', columnIndex, 'Headers length:', headers.length);
    return;
  }
  
  // Update node with column info
  node.columnIndex = columnIndex;
  node.columnName = headers[columnIndex] || `Column ${columnIndex + 1}`;
  
  // Get contextually filtered values
  const contextualValues = getContextualValuesForNode(node, columnIndex);
  
  if (contextualValues.length === 0) {
    // 🔧 FIX: No values found → LEAVE EMPTY, don't use column name!
    node.value = '';
    applyTemplateToSiblingsIfFirst(node);
    if (window.renderSheet) {
      window.renderSheet(activeSheetId, sheetData);
    }
    return;
  }
  
  // Create dropdown container with checkboxes
  const valueContainer = document.createElement('div');
  valueContainer.style.position = 'relative';
  valueContainer.style.display = 'inline-block';
  
  const valueSelect = document.createElement('select');
  valueSelect.className = 'value-selector';
  valueSelect.style.padding = '2px 6px';
  valueSelect.style.fontSize = '12px';
  valueSelect.style.border = '2px solid #4caf50';
  valueSelect.style.borderRadius = '3px';
  valueSelect.style.background = '#e8f5e9';
  valueSelect.style.color = '#4caf50';
  valueSelect.style.minWidth = '150px';
  valueSelect.multiple = true;
  valueSelect.size = Math.min(contextualValues.length + 1, 6); // Show max 6 items
  
  // Add instruction option
  const instructionOption = document.createElement('option');
  instructionOption.disabled = true;
  instructionOption.textContent = `✓ Select ${node.columnName} values:`;
  instructionOption.style.fontWeight = 'bold';
  instructionOption.style.background = '#f0f0f0';
  valueSelect.appendChild(instructionOption);
  
  // Add value options
  contextualValues.forEach(value => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    valueSelect.appendChild(option);
  });
  
  // Add apply button
  const applyBtn = document.createElement('button');
  applyBtn.textContent = 'Apply';
  applyBtn.style.marginLeft = '5px';
  applyBtn.style.padding = '2px 8px';
  applyBtn.style.fontSize = '12px';
  applyBtn.style.border = '1px solid #4caf50';
  applyBtn.style.borderRadius = '3px';
  applyBtn.style.background = '#4caf50';
  applyBtn.style.color = 'white';
  applyBtn.style.cursor = 'pointer';
  
  applyBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    const selectedValues = Array.from(valueSelect.selectedOptions).map(opt => opt.value);
    
    if (selectedValues.length === 0) {
      alert('Please select at least one value');
      return;
    }
    
    // Apply selected values
    applyColumnValues(node, selectedValues);
  });
  
  valueContainer.appendChild(valueSelect);
  valueContainer.appendChild(applyBtn);
  
  // Replace the old select with the new value selector
  oldSelect.replaceWith(valueContainer);
}

// Helper function to get raw Excel data
function getRawExcelData(sheetId) {
  // Try to get from window.rawExcelData first
  if (window.rawExcelData && window.rawExcelData[sheetId]) {
    return window.rawExcelData[sheetId];
  }
  
  // Try direct array access
  if (window.rawExcelData && Array.isArray(window.rawExcelData)) {
    return window.rawExcelData;
  }
  
  // Fallback to sheet data
  const sheetData = window.excelData.sheetsLoaded[sheetId];
  if (sheetData && sheetData.rawData) {
    return sheetData.rawData;
  }
  
  return null;
}



// Helper function to find the template parent (first sibling)
function findTemplateParent(currentParent, root) {
  // Find the parent of currentParent
  const grandParent = findParentInStructure(root, currentParent.id);
  if (!grandParent || !grandParent.children || grandParent.children.length === 0) return null;
  
  // Return the first child (template)
  return grandParent.children[0];
}

// Function to show multicolumn selector with checkboxes
function showMultiColumnSelector(node) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  if (!sheetData || !sheetData.headers) {
    alert('No sheet data available');
    return;
  }
  
  // Create modal
  const modal = document.createElement('div');
  modal.className = 'pdf-preview-modal';
  modal.style.zIndex = '10000';
  
  const content = document.createElement('div');
  content.className = 'pdf-preview-content';
  content.style.maxWidth = '500px';
  content.style.maxHeight = '80vh';
  content.style.overflow = 'auto';
  
  // Close button
  const closeBtn = document.createElement('button');
  closeBtn.className = 'pdf-preview-close';
  closeBtn.innerHTML = '×';
  closeBtn.addEventListener('click', () => {
    document.body.removeChild(modal);
  });
  
  // Header
  const header = document.createElement('h2');
  header.textContent = 'Select Multiple Columns';
  header.style.marginBottom = '20px';
  header.style.color = '#34a3d7';
  
  // Description
  const description = document.createElement('p');
  description.textContent = 'Select multiple columns to create child containers for each selected column:';
  description.style.marginBottom = '20px';
  description.style.color = '#666';
  
  // Get used columns to filter them out
  const usedColumns = getUsedColumnsInHierarchy(node);
  
  // Checkbox container
  const checkboxContainer = document.createElement('div');
  checkboxContainer.style.maxHeight = '300px';
  checkboxContainer.style.overflowY = 'auto';
  checkboxContainer.style.border = '1px solid #ddd';
  checkboxContainer.style.borderRadius = '4px';
  checkboxContainer.style.padding = '10px';
  checkboxContainer.style.marginBottom = '20px';
  
  const checkboxes = [];
  
  // Create checkboxes for each available column
  sheetData.headers.forEach((header, index) => {
    if (!header) return;
    
    const checkboxWrapper = document.createElement('div');
    checkboxWrapper.style.display = 'flex';
    checkboxWrapper.style.alignItems = 'center';
    checkboxWrapper.style.padding = '8px';
    checkboxWrapper.style.borderBottom = '1px solid #f0f0f0';
    checkboxWrapper.style.cursor = 'pointer';
    
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.id = `column-${index}`;
    checkbox.value = index;
    checkbox.style.marginRight = '10px';
    
    const label = document.createElement('label');
    label.htmlFor = `column-${index}`;
    label.style.cursor = 'pointer';
    label.style.flex = '1';
    
    // Check if column is already used
    if (usedColumns.includes(index)) {
      label.textContent = `${header} (already used)`;
      label.style.color = '#999';
      label.style.fontStyle = 'italic';
      checkbox.disabled = true;
      checkboxWrapper.style.opacity = '0.5';
    } else {
      label.textContent = header;
      label.style.color = '#333';
    }
    
    // Add hover effect for enabled items
    if (!checkbox.disabled) {
      checkboxWrapper.addEventListener('mouseenter', () => {
        checkboxWrapper.style.backgroundColor = '#f8f9fa';
      });
      
      checkboxWrapper.addEventListener('mouseleave', () => {
        checkboxWrapper.style.backgroundColor = 'transparent';
      });
      
      // Allow clicking on wrapper to toggle checkbox
      checkboxWrapper.addEventListener('click', (e) => {
        if (e.target !== checkbox) {
          checkbox.checked = !checkbox.checked;
        }
      });
    }
    
    checkboxWrapper.appendChild(checkbox);
    checkboxWrapper.appendChild(label);
    checkboxContainer.appendChild(checkboxWrapper);
    
    if (!checkbox.disabled) {
      checkboxes.push(checkbox);
    }
  });
  
  // Select All / Deselect All buttons
  const selectButtonsContainer = document.createElement('div');
  selectButtonsContainer.style.display = 'flex';
  selectButtonsContainer.style.gap = '10px';
  selectButtonsContainer.style.marginBottom = '20px';
  
  const selectAllBtn = document.createElement('button');
  selectAllBtn.textContent = 'Select All';
  selectAllBtn.className = 'control-button';
  selectAllBtn.style.fontSize = '12px';
  selectAllBtn.style.padding = '6px 12px';
  selectAllBtn.addEventListener('click', () => {
    checkboxes.forEach(cb => cb.checked = true);
  });
  
  const deselectAllBtn = document.createElement('button');
  deselectAllBtn.textContent = 'Deselect All';
  deselectAllBtn.className = 'control-button';
  deselectAllBtn.style.fontSize = '12px';
  deselectAllBtn.style.padding = '6px 12px';
  deselectAllBtn.addEventListener('click', () => {
    checkboxes.forEach(cb => cb.checked = false);
  });
  
  selectButtonsContainer.appendChild(selectAllBtn);
  selectButtonsContainer.appendChild(deselectAllBtn);
  
  // Action buttons
  const buttonContainer = document.createElement('div');
  buttonContainer.style.display = 'flex';
  buttonContainer.style.justifyContent = 'space-between';
  buttonContainer.style.gap = '10px';
  
  const cancelBtn = document.createElement('button');
  cancelBtn.textContent = 'Cancel';
  cancelBtn.className = 'control-button';
  cancelBtn.style.backgroundColor = '#6c757d';
  cancelBtn.style.color = 'white';
  
  const applyBtn = document.createElement('button');
  applyBtn.textContent = 'Create Containers';
  applyBtn.className = 'control-button';
  applyBtn.style.backgroundColor = '#4caf50';
  applyBtn.style.color = 'white';
  
  buttonContainer.appendChild(cancelBtn);
  buttonContainer.appendChild(applyBtn);
  
  // Event listeners
  cancelBtn.addEventListener('click', () => {
    document.body.removeChild(modal);
  });
  
  applyBtn.addEventListener('click', () => {
    // 🔍 DEBUG: Log checkboxes and their values
    console.log('🔍 DEBUG: Total checkboxes available:', checkboxes.length);
    console.log('🔍 DEBUG: All checkboxes:', checkboxes.map(cb => ({
      value: cb.value, 
      checked: cb.checked, 
      disabled: cb.disabled,
      name: sheetData.headers[parseInt(cb.value)]
    })));
    
    const selectedColumns = checkboxes
      .filter(cb => cb.checked)
      .map(cb => ({
        index: parseInt(cb.value),
        name: sheetData.headers[parseInt(cb.value)]
      }));
    
    console.log('🔍 DEBUG: Selected columns after filtering:', selectedColumns);
    
    if (selectedColumns.length === 0) {
      alert('Please select at least one column');
      return;
    }
    
    // Close modal
    document.body.removeChild(modal);
    
    // Create child containers for each selected column
    createMultipleColumnContainers(node, selectedColumns);
  });
  
  // Assemble modal
  content.appendChild(closeBtn);
  content.appendChild(header);
  content.appendChild(description);
  content.appendChild(selectButtonsContainer);
  content.appendChild(checkboxContainer);
  content.appendChild(buttonContainer);
  modal.appendChild(content);
  document.body.appendChild(modal);
  
  // Focus on modal for keyboard navigation
  modal.focus();
}

// Function to create multiple column containers
function createMultipleColumnContainers(parentNode, selectedColumns) {
  // Validate inputs
  if (!parentNode) {
    console.error('❌ createMultipleColumnContainers: parentNode is null or undefined');
    return;
  }
  
  if (!selectedColumns || !Array.isArray(selectedColumns) || selectedColumns.length === 0) {
    console.error('❌ createMultipleColumnContainers: selectedColumns is invalid or empty');
    return;
  }
  
  // If this is a "New Container" without column info, 
  // find the real template parent (Les node) to add children to
  let realTemplateParent = parentNode;
  let shouldRemoveContainer = false;
  
  if (!parentNode.columnIndex && parentNode.columnIndex !== 0 && parentNode.parent) {
    realTemplateParent = parentNode.parent;
    shouldRemoveContainer = true; // Mark for removal since we're bypassing it
  }
  
  // Additional safety check
  if (!realTemplateParent) {
    console.error('❌ realTemplateParent is null, using original parentNode');
    realTemplateParent = parentNode;
    shouldRemoveContainer = false;
  }
  
  // Remove placeholder if it exists
  if (realTemplateParent.isPlaceholder) {
    realTemplateParent.isPlaceholder = false;
    realTemplateParent.value = 'Container';
  }
  
  // Ensure children array exists and properties array exists
  if (!realTemplateParent.children) {
    realTemplateParent.children = [];
  }
  if (!realTemplateParent.properties) {
    realTemplateParent.properties = [];
  }
  
  // Create a child container for each selected column
  selectedColumns.forEach((column, index) => {
    // Validate column before processing
    if (!column || column.index === undefined || column.index === null || isNaN(column.index)) {
      console.error('❌ Invalid column in selectedColumns:', column);
      return; // Skip this invalid column
    }
    
    const newNode = {
      id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
      value: '',  // Start with empty value, not column name
      columnName: column.name,
      columnIndex: column.index,
      level: realTemplateParent.level + 1,
      children: [],
      properties: [],
      layoutMode: 'horizontal',
      parent: realTemplateParent  // Parent is now the real template!
    };
    
    realTemplateParent.children.push(newNode);
    
    // FIXED: Get single specific value instead of all unique values
    setTimeout(() => {
      assignSpecificValueToNode(newNode, column.index);
    }, 100);
  });
  
  // 🔄 CRUCIAL: If parent is a template, replicate new children to all duplicates!
  setTimeout(() => {
    // 🗑️ Remove the "New Container" if we bypassed it
    if (shouldRemoveContainer && realTemplateParent.children) {
      console.log('🗑️ Removing empty "New Container" since we added children directly to template');
      const containerIndex = realTemplateParent.children.findIndex(child => child.id === parentNode.id);
      if (containerIndex !== -1) {
        realTemplateParent.children.splice(containerIndex, 1);
      }
    }
    
    console.log('🔍 DEBUG: Template parent node details:', {
      value: realTemplateParent.value,
      isTemplate: realTemplateParent.isTemplate,
      isDuplicate: realTemplateParent.isDuplicate,
      hasParent: !!realTemplateParent.parent,
      parentValue: realTemplateParent.parent ? realTemplateParent.parent.value : 'NO PARENT'
    });
    
    // Get root to do global template marking
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    const rootNode = sheetData.root;
    
    // Mark real template parent as GLOBAL template if needed
    // Add null check for realTemplateParent
    if (realTemplateParent && realTemplateParent.columnName !== undefined) {
      markGlobalTemplatesOfType(rootNode, realTemplateParent.columnName, realTemplateParent.columnIndex);
    } else {
      console.warn('⚠️ realTemplateParent is null or missing columnName, skipping template marking');
    }
    
    console.log('🔍 DEBUG: After markGlobalTemplatesOfType:', {
      isTemplate: realTemplateParent.isTemplate,
      isDuplicate: realTemplateParent.isDuplicate
    });
    
    // Apply template replication if real parent is template
    if (realTemplateParent.isTemplate) {
      console.log('🔄 Real template parent - replicating new children to ALL nodes of same type');
      applyGlobalTemplateReplication(realTemplateParent);
    } else {
      console.log('⚠️ Real parent is NOT template - no replication');
      // Re-render with proper styling
      const activeSheetId = window.excelData.activeSheetId;
      const sheetData = window.excelData.sheetsLoaded[activeSheetId];
      
      if (window.renderSheet) {
        window.renderSheet(activeSheetId, sheetData);
      }
    }
  }, 200); // Wait a bit longer for all children to be processed
  
  // Apply flexbox styling after render
  applyFlexboxStyling();
}

// Function to toggle node layout (IMPROVED: No full re-render)
function toggleNodeLayout(node) {
  // Check if this is a duplicate - only templates can be modified
  if (node.isDuplicate) {
    alert('🚫 This is a duplicate node. Please modify the template (- master) instead.');
    return;
  }
  
  // Toggle layout mode
  const currentLayout = node.layoutMode || 'horizontal';
  node.layoutMode = currentLayout === 'horizontal' ? 'vertical' : 'horizontal';
  
  console.log(`🔄 Layout toggled to ${node.layoutMode} for ${node.value}`);
  
  // Find the node element in the DOM
  const nodeElement = document.querySelector(`[data-node-id="${node.id}"]`);
  if (!nodeElement) {
    console.error('Could not find node element for layout toggle');
    return;
  }
  
  // Find the children container within this node - ALL LEVELS
  let childrenContainer = nodeElement.querySelector('.node-children, .level-0-children, .level-1-children, .level-2-children, .level-3-children, .level-4-children, .level-5-children, .level-6-children, .level-7-children, .level-8-children, .level-9-children, .level-10-children');
  
  // If no children container exists, create one for future use
  if (!childrenContainer) {
    console.log(`Creating children container for layout toggle on ${node.value}`);
    
    // Create children container based on node level
    const nodeLevel = Math.abs(node.level || 0);
    childrenContainer = document.createElement('div');
    childrenContainer.className = `level-${nodeLevel}-children`;
    
    // Find the content area to append the children container
    const contentArea = nodeElement.querySelector(`.level-${nodeLevel}-content, .node-content`);
    if (contentArea) {
      contentArea.appendChild(childrenContainer);
    } else {
      // Fallback: append to the node element itself
      nodeElement.appendChild(childrenContainer);
    }
    
    // Initialize children array if it doesn't exist
    if (!node.children) {
      node.children = [];
    }
  }
  
  // This logic is now handled below, so remove this duplicate
  
  // Remove ALL layout classes - both old and new
  childrenContainer.classList.remove('layout-horizontal', 'layout-vertical');
  // Remove old level-specific layout classes
  for (let i = 0; i <= 10; i++) {
    childrenContainer.classList.remove(`level-${i}-horizontal-layout`);
    childrenContainer.classList.remove(`level-${i}-vertical-layout`);
  }
  
  // Add only the new simple layout class
  if (node.layoutMode === 'vertical') {
    childrenContainer.classList.add('layout-vertical');
  } else {
    childrenContainer.classList.add('layout-horizontal');
  }
  
  console.log(`✅ Layout updated visually to ${node.layoutMode} for ${node.value}`);
  
  // If this is a template, apply the same layout change to ALL nodes of same type
  if (node.isTemplate) {
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    const rootNode = sheetData.root;
    const allNodesOfType = findAllNodesOfType(rootNode, node.columnName, node.columnIndex);
    const duplicates = allNodesOfType.filter(n => n !== node);
    
    duplicates.forEach(duplicate => {
      duplicate.layoutMode = node.layoutMode;
      
      // Update the duplicate's visual layout too
      const duplicateElement = document.querySelector(`[data-node-id="${duplicate.id}"]`);
      if (duplicateElement) {
        const duplicateChildrenContainer = duplicateElement.querySelector('.node-children, .level-0-children, .level-1-children, .level-2-children, .level-3-children');
        if (duplicateChildrenContainer) {
          duplicateChildrenContainer.classList.remove('layout-horizontal', 'layout-vertical');
          duplicateChildrenContainer.classList.add(`layout-${duplicate.layoutMode}`);
          
          if (duplicate.layoutMode === 'vertical') {
            duplicateChildrenContainer.classList.remove('level-0-horizontal-layout', 'level-1-horizontal-layout', 'level-2-horizontal-layout');
            duplicateChildrenContainer.classList.add(`level-${duplicate.level || 0}-vertical-layout`);
          } else {
            duplicateChildrenContainer.classList.remove('level-0-vertical-layout', 'level-1-vertical-layout', 'level-2-vertical-layout');
            duplicateChildrenContainer.classList.add(`level-${duplicate.level || 0}-horizontal-layout`);
          }
        }
      }
    });
    
    console.log(`✅ Layout updated for ${duplicates.length} duplicates`);
  }
}

// Function to apply layout styling to all nodes
function applyLayoutStyling() {
  // Apply CSS classes based on layout mode
  document.querySelectorAll('[data-node-id]').forEach(nodeEl => {
    const nodeId = nodeEl.getAttribute('data-node-id');
    const node = findNodeById(nodeId);
    
    if (node && node.layoutMode) {
      const childrenContainer = nodeEl.querySelector('.node-children, .level-0-children, .level-1-children, .level-2-children, .level-3-children');
      if (childrenContainer) {
        // Remove existing layout classes
        childrenContainer.classList.remove('layout-horizontal', 'layout-vertical');
        
        // Add new layout class
        childrenContainer.classList.add(`layout-${node.layoutMode}`);
      }
    }
  });
}

// Helper function to find a node's parent in the hierarchy
function findParentInStructure(root, nodeId) {
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
  
  return searchChildren(root);
}

// Export functions to global scope
window.enableLiveEditing = enableLiveEditing;
window.addEditControls = addEditControls;
window.applyFlexboxStyling = applyFlexboxStyling;
window.applyLayoutStyling = applyLayoutStyling;
window.toggleNodeLayout = toggleNodeLayout;
window.addChildContainer = addChildContainer;
window.addSiblingContainer = addSiblingContainer;
window.deleteContainer = deleteContainer;
window.assignColumnToNode = assignColumnToNode;
window.assignSpecificValueToNode = assignSpecificValueToNode;
window.findParentInStructure = findParentInStructure;
window.markGlobalTemplatesOfType = markGlobalTemplatesOfType;
window.applyGlobalTemplateReplication = applyGlobalTemplateReplication;
window.findAllNodesOfType = findAllNodesOfType;

// Export as module for structure-builder to avoid conflicts
window.ExcelViewerLiveEditor = {
  addChildContainer,
  addSiblingContainer,
  deleteContainer,
  assignColumnToNode,
  assignSpecificValueToNode
};

// SIMPLE template replication - when template changes, copy to all duplicates
function applySimpleTemplateReplication(templateNode) {
  console.log('🔄 Simple template replication for:', templateNode.value);
  
  // Find all duplicates that reference this template
  const duplicates = findDuplicatesOfTemplate(templateNode.id);
  
  console.log(`Found ${duplicates.length} duplicates to update`);
  
  // Copy template structure to all duplicates
  duplicates.forEach(duplicate => {
    console.log(`Copying structure to duplicate: ${duplicate.value}`);
    
    // Clone the template's children structure
    duplicate.children = cloneStructure(templateNode.children);
    duplicate.layoutMode = templateNode.layoutMode;
    
    // Populate with contextual data for this duplicate
    populateWithSimpleData(duplicate);
  });
  
  // Re-render
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  if (window.renderSheet) {
    window.renderSheet(activeSheetId, sheetData);
  }
  applyFlexboxStyling();
}

// Find all duplicates of a template (simple version)
function findDuplicatesOfTemplate(templateId) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const duplicates = [];
  
  function searchForDuplicates(node) {
    if (node.templateId === templateId && node.isDuplicate) {
      duplicates.push(node);
    }
    if (node.children) {
      node.children.forEach(searchForDuplicates);
    }
  }
  
  searchForDuplicates(sheetData.root);
  return duplicates;
}

// Clone structure with new IDs (simple version)
function cloneStructure(children) {
  if (!children || children.length === 0) return [];
  
  return children.map(child => {
    console.log(`   🔧 Cloning child: "${child.value}" with columnIndex: ${child.columnIndex}, columnName: "${child.columnName}"`);
    return {
      id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
      value: child.columnName || 'New Container', // Reset value - will be populated contextually
      columnName: child.columnName,
      columnIndex: child.columnIndex,
      level: child.level,
      children: cloneStructure(child.children),
      properties: child.properties ? [...child.properties] : [],
      layoutMode: child.layoutMode || 'horizontal'
    };
  });
}

// Simple data population - just assign columns to children
function populateWithSimpleData(duplicateNode) {
  // For each child that has a column, just call assignColumnToNode
  if (duplicateNode.children && duplicateNode.children.length > 0) {
    duplicateNode.children.forEach((child, index) => {
      if (child.columnIndex !== undefined && child.columnIndex !== null) {
        // Call assignColumnToNode with skipReplication=true to avoid infinite loops
        assignColumnToNode(child, child.columnIndex, true);
      }
    });
  }
} 

// GLOBAL: Helper function to find ALL nodes of the same type across the entire tree
function findAllNodesOfType(rootNode, columnName, columnIndex) {
  const allNodes = [];
  
  function traverse(node) {
    if (node.columnName === columnName && node.columnIndex === columnIndex) {
      allNodes.push(node);
    }
    if (node.children && node.children.length > 0) {
      node.children.forEach(traverse);
    }
  }
  
  traverse(rootNode);
  return allNodes;
}

// GLOBAL: Mark templates and duplicates based on TYPE (columnName + columnIndex) across entire tree
function markGlobalTemplatesOfType(rootNode, columnName, columnIndex) {
  const allNodesOfType = findAllNodesOfType(rootNode, columnName, columnIndex);
  
  if (allNodesOfType.length <= 1) return; // No templates needed for single nodes
  
  // First node becomes the template
  const template = allNodesOfType[0];
  template.isTemplate = true;
  template.isDuplicate = false;
  console.log(`- master GLOBAL: Marked "${template.value}" (${columnName}) as GLOBAL template`);
  
  // All others become duplicates
  allNodesOfType.slice(1).forEach(duplicate => {
    duplicate.isDuplicate = true;
    duplicate.isTemplate = false;
    duplicate.templateId = template.id;
    console.log(`📋 GLOBAL: Marked "${duplicate.value}" (${columnName}) as GLOBAL duplicate`);
  });
}

// Helper function to mark existing nodes as complete if they have column assignments
function markExistingNodesAsComplete(rootNode) {
  if (!rootNode) return;
  
  // If node has columnIndex but no isIncomplete flag, it's complete
  if (rootNode.columnIndex !== undefined && rootNode.columnIndex !== null && rootNode.isIncomplete === undefined) {
    rootNode.isIncomplete = false;
  }
  
  // If node has no columnIndex and no isIncomplete flag, it's incomplete
  if ((rootNode.columnIndex === undefined || rootNode.columnIndex === null) && rootNode.isIncomplete === undefined) {
    rootNode.isIncomplete = true;
  }
  
  // Recursively process children
  if (rootNode.children && rootNode.children.length > 0) {
    rootNode.children.forEach(child => markExistingNodesAsComplete(child));
  }
}

// 🔧 GLOBAL TEMPLATE REPLICATION: Copy STRUCTURE to ALL nodes of same TYPE across entire tree
function applyGlobalTemplateReplication(templateNode) {
  console.log('🔧 GLOBAL template replication for:', templateNode.value, templateNode.columnName);
  
  if (!templateNode.isTemplate) {
    console.log('❌ Node is not marked as template, skipping replication');
    return;
  }
  
  // Get root node to search entire tree
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const rootNode = sheetData.root;
  
  // Find ALL nodes of the same type (columnName + columnIndex) across entire tree
  const allNodesOfType = findAllNodesOfType(rootNode, templateNode.columnName, templateNode.columnIndex);
  const duplicates = allNodesOfType.filter(node => node !== templateNode);
  
  console.log(`🌍 GLOBAL: Found ${duplicates.length} duplicates across entire tree to replicate structure to`);
  
  // For each duplicate: copy the ACTIONS but with their own Excel context
  duplicates.forEach(duplicate => {
    console.log(`🔧 GLOBAL: Replicating structure to duplicate: ${duplicate.value} in ${getNodePath(duplicate)}`);
    
    // Copy the template's children STRUCTURE (columnIndex/columnName assignments)
    templateNode.children.forEach(templateChild => {
      console.log(`  📋 GLOBAL: Copying action: assign column ${templateChild.columnName} (${templateChild.columnIndex})`);
      
      // Skip nodes that don't have column assignments yet
      if (templateChild.columnIndex === undefined || templateChild.columnIndex === null) {
        console.log(`  ⏭️ GLOBAL: Skipping child without column assignment: ${templateChild.value}`);
        return;
      }
      
      // Check if this duplicate already has this column assigned
      const existingChild = duplicate.children.find(child => 
        child.columnIndex === templateChild.columnIndex
      );
      
      if (!existingChild) {
        // Apply the same column assignment action to this duplicate
        // But let it use its own Excel context to determine what values to create
        console.log(`  ➕ GLOBAL: Applying column assignment to duplicate "${duplicate.value}"`);
        
        // Temporarily create a placeholder child and assign the column
        const placeholderChild = {
          id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
          value: 'New Container',
          columnName: '',
          level: duplicate.level + 1,
          children: [],
          properties: [],
          layoutMode: 'horizontal',
          parent: duplicate
        };
        
        duplicate.children.push(placeholderChild);
        
        // Now assign the column - this will use the duplicate's context to get the right Excel values
        assignColumnToNode(placeholderChild, templateChild.columnIndex, true);
      } else {
        console.log(`  ⏭️ GLOBAL: Duplicate already has column ${templateChild.columnName} assigned`);
      }
      
      // ALSO copy any existing children structure from template child to duplicate child
      if (templateChild.children && templateChild.children.length > 0) {
        const targetChild = duplicate.children.find(child => 
          child.columnIndex === templateChild.columnIndex
        );
        if (targetChild && (!targetChild.children || targetChild.children.length === 0)) {
          console.log(`  🔄 GLOBAL: Copying ${templateChild.children.length} children from template to duplicate`);
          targetChild.children = JSON.parse(JSON.stringify(templateChild.children));
          // Update parent references
          targetChild.children.forEach(grandchild => {
            grandchild.parent = targetChild;
          });
        }
      }
    });
  });
  
  // Re-render to show the changes
  if (window.ExcelViewerSheetManager && window.ExcelViewerSheetManager.renderSheet) {
    window.ExcelViewerSheetManager.renderSheet(activeSheetId, sheetData);
  } else if (window.renderSheet) {
    window.renderSheet(activeSheetId, sheetData);
  }
  applyFlexboxStyling();
}

// Helper function to get node path for debugging
function getNodePath(node) {
  const path = [];
  let current = node;
  while (current && current.parent) {
    if (current.value && current.columnName) {
      path.unshift(`${current.columnName}:${current.value}`);
    }
    current = current.parent;
  }
  return path.join(' → ');
}

// 🚫 REMOVED: No more template cloning - each node reads Excel directly

// REMOVED: Old MATRIX LOOKUP template copying is gone - use smart replication instead
function populateWithContextualData(templateNode, duplicateNode) {
  // NO-OP: Smart replication (applySmartTemplateReplication) handles structure copying with contextual Excel data
}

// Helper function: Find Excel row that matches exact path coordinates
function findExcelRowByPath(excelPath) {
  // Use the same logic as getRawExcelData function 
  const activeSheetId = window.excelData.activeSheetId;
  let rawData = null;
  
  // Try to get from window.rawExcelData first
  if (window.rawExcelData && window.rawExcelData[activeSheetId]) {
    rawData = window.rawExcelData[activeSheetId];
  } else if (window.rawExcelData && Array.isArray(window.rawExcelData)) {
    rawData = window.rawExcelData;
  } else {
    // Fallback to sheet data
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    if (sheetData && sheetData.rawData) {
      rawData = sheetData.rawData;
    }
  }
  
  if (!rawData || !rawData.length) {
    return null;
  }

  // Skip header row, check each data row
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    let isMatch = true;

    // Check if this row matches ALL path coordinates
    for (const pathItem of excelPath) {
      const cellValue = row[pathItem.columnIndex];
      const cellString = cellValue ? cellValue.toString().trim() : '';
      const matches = cellString === pathItem.value;
      
      if (!matches) {
        isMatch = false;
        break;
      }
    }

    if (isMatch) {
      return row;
    }
  }

  return null;
} 

// FIXED: Template replication with correct filtering logic
function getContextualValuesForNodeWithFilter(filterNode, columnIndex) {
  const activeSheetId = window.excelData.activeSheetId;
  const rawData = getRawExcelData(activeSheetId);
  
  if (!rawData || rawData.length <= 1) {
    return [];
  }
  
  // Get ONLY parent context (don't include filter node)
  const parentContext = getFullParentContext(filterNode);
  
  // First find the range of rows for parent context
  let rowRange = [];
  if (parentContext.length === 0) {
    // No parent context - use all rows
    rowRange = rawData.slice(1); // Skip header
  } else {
    // Find rows matching parent hierarchy
    rowRange = findRowsForParentHierarchy(rawData, parentContext);
  }
  
  // Within that range, filter for the specific filter value
  const filterColumnIdx = filterNode.columnIndex;
  const filterValue = filterNode.value;
  
  const matchingRows = rowRange.filter(row => {
    const cellValue = row[filterColumnIdx];
    const matches = cellValue && cellValue.toString().trim() === filterValue;
    return matches;
  });
  
  if (matchingRows.length === 0) {
    return [];
  }
  
  // Extract unique values from the specific column within filtered rows
  const uniqueValues = new Set();
  matchingRows.forEach((row, index) => {
    const value = row[columnIndex];
    if (value && value.toString().trim()) {
      uniqueValues.add(value.toString().trim());
    }
  });
  
  const result = Array.from(uniqueValues).sort();
  
  return result;
} 

// Store current selection state
let currentSelectedNode = null;
let currentSelectionMode = 'child'; // 'child' or 'sibling'
let selectedColumnIndices = []; // Changed to array for multiselect
let currentSelectionArea = null;

// Function to show the column selection area
function showColumnSelectionArea(node) {
  console.log('showColumnSelectionArea called with node:', node);
  
  // Hide any existing selection area
  hideColumnSelectionArea();
  
  currentSelectedNode = node;
  selectedColumnIndices = []; // Reset selections
  
  // Create floating column selection area
  const columnSelectionArea = document.createElement('div');
  columnSelectionArea.className = 'column-selection-area';
  columnSelectionArea.style.display = 'block';
  
  // Find the node element to position relative to
  const nodeElement = document.querySelector(`[data-node-id="${node.id}"]`);
  if (!nodeElement) {
    console.error('Could not find node element for positioning');
    return;
  }
  
  // Position the area below the node and match its width
  const nodeRect = nodeElement.getBoundingClientRect();
  columnSelectionArea.style.top = (nodeRect.bottom + window.scrollY + 10) + 'px';
  columnSelectionArea.style.left = (nodeRect.left + window.scrollX) + 'px';
  columnSelectionArea.style.width = nodeRect.width + 'px';
  columnSelectionArea.style.minWidth = 'auto'; // Override the CSS min-width
  columnSelectionArea.style.maxWidth = 'none'; // Override the CSS max-width
  
  // Create the content structure
  columnSelectionArea.innerHTML = `
    <div class="column-controls">
      <div class="toggle-group">
        <label>Add as:</label>
        <button class="toggle-button active child-toggle" data-mode="child">Child</button>
        <button class="toggle-button sibling-toggle" data-mode="sibling">Sibling</button>
      </div>
      <div class="available-columns">
        <!-- Column buttons will be added here -->
      </div>
      <div class="selection-actions" style="display: none;">
        <span class="selection-count">0 columns selected</span>
        <button class="add-button">➕ Add Selected</button>
      </div>
    </div>
  `;
  
  // Add to body
  document.body.appendChild(columnSelectionArea);
  currentSelectionArea = columnSelectionArea;
  
  // Get containers
  const availableColumnsContainer = columnSelectionArea.querySelector('.available-columns');
  
  // Get available columns
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  if (!sheetData || !sheetData.headers) {
    availableColumnsContainer.innerHTML = '<span style="color: #999;">No columns available</span>';
    return;
  }
  
  const usedColumns = getUsedColumnsInHierarchy(node);
  
  // Create column buttons
  sheetData.headers.forEach((header, index) => {
    if (!header) return;
    
    const columnBtn = document.createElement('button');
    columnBtn.className = 'column-button';
    columnBtn.textContent = header;
    columnBtn.dataset.columnIndex = index;
    
    // Add tooltip for long text
    if (header.length > 15) {
      columnBtn.title = header;
    }
    
    if (usedColumns.includes(index)) {
      columnBtn.classList.add('disabled');
      columnBtn.title = `${header} (already used in hierarchy)`;
      columnBtn.disabled = true;
    } else {
      columnBtn.addEventListener('click', () => toggleColumnSelection(index, columnBtn));
    }
    
    availableColumnsContainer.appendChild(columnBtn);
  });
  
  // Setup toggle buttons for this area
  setupToggleButtonsForArea(columnSelectionArea);
  
  // Setup add button for this area
  setupAddButtonForArea(columnSelectionArea);
}

// Function to hide column selection area
function hideColumnSelectionArea() {
  if (currentSelectionArea) {
    currentSelectionArea.remove();
    currentSelectionArea = null;
  }
  currentSelectedNode = null;
  selectedColumnIndices = [];
}

// Function to setup toggle buttons for specific area
function setupToggleButtonsForArea(area) {
  const childToggle = area.querySelector('.child-toggle');
  const siblingToggle = area.querySelector('.sibling-toggle');
  
  childToggle.addEventListener('click', () => {
    currentSelectionMode = 'child';
    childToggle.classList.add('active');
    siblingToggle.classList.remove('active');
  });
  
  siblingToggle.addEventListener('click', () => {
    currentSelectionMode = 'sibling';
    siblingToggle.classList.add('active');
    childToggle.classList.remove('active');
  });
  
  // Set initial state
  if (currentSelectedNode && (currentSelectedNode.level <= -1 || currentSelectedNode.isPlaceholder)) {
    // Root level nodes can only have children
    siblingToggle.disabled = true;
    siblingToggle.style.opacity = '0.5';
    currentSelectionMode = 'child';
    childToggle.classList.add('active');
    siblingToggle.classList.remove('active');
  } else {
    siblingToggle.disabled = false;
    siblingToggle.style.opacity = '1';
    // Default to current mode
    if (currentSelectionMode === 'child') {
      childToggle.classList.add('active');
      siblingToggle.classList.remove('active');
    } else {
      siblingToggle.classList.add('active');
      childToggle.classList.remove('active');
    }
  }
}

// Function to toggle column selection (multiselect)
function toggleColumnSelection(columnIndex, buttonElement) {
  const isSelected = buttonElement.classList.contains('selected');
  
  if (isSelected) {
    // Deselect this column
    buttonElement.classList.remove('selected');
    const index = selectedColumnIndices.indexOf(columnIndex);
    if (index > -1) {
      selectedColumnIndices.splice(index, 1);
    }
  } else {
    // Select this column
    buttonElement.classList.add('selected');
    selectedColumnIndices.push(columnIndex);
  }
  
  // Update selection UI
  updateSelectionUI();
}

// Function to update selection UI
function updateSelectionUI() {
  if (!currentSelectionArea) return;
  
  const selectionActions = currentSelectionArea.querySelector('.selection-actions');
  const selectionCount = currentSelectionArea.querySelector('.selection-count');
  const addButton = currentSelectionArea.querySelector('.add-button');
  
  if (selectedColumnIndices.length === 0) {
    selectionActions.style.display = 'none';
  } else {
    selectionActions.style.display = 'flex';
    selectionActions.style.alignItems = 'center';
    selectionActions.style.gap = '10px';
    
    const count = selectedColumnIndices.length;
    const columnText = count === 1 ? 'column' : 'columns';
    selectionCount.textContent = `${count} ${columnText} selected`;
  }
}

// Function to setup add button for specific area
function setupAddButtonForArea(area) {
  const addButton = area.querySelector('.add-button');
  
  addButton.addEventListener('click', () => {
    if (!currentSelectedNode) {
      console.error('❌ No currentSelectedNode available');
      return;
    }
    
    if (selectedColumnIndices.length === 0) {
      console.error('❌ No columns selected');
      return;
    }
    
    // 🔥 CRITICAL FIX: Save selections BEFORE hiding (which resets them!)
    const savedSelectedIndices = [...selectedColumnIndices]; // Create copy
    const savedSelectionMode = currentSelectionMode;
    const savedSelectedNode = currentSelectedNode;
    
    // Hide the selection area (this resets selectedColumnIndices!)
    hideColumnSelectionArea();
    
    // Apply the selections using SAVED values
    
    if (savedSelectedIndices.length === 1) {
      // Single column selection
      if (savedSelectionMode === 'child') {
        // 🔥 CRITICAL FIX: Create child first, then assign column to the child
        addChildContainer(savedSelectedNode, false); // Don't skip re-render for single child
        
        // Find the newly created child (it will be the last child) - DO IT SYNCHRONOUSLY
        if (savedSelectedNode.children && savedSelectedNode.children.length > 0) {
          const newChild = savedSelectedNode.children[savedSelectedNode.children.length - 1];
          assignColumnToNode(newChild, savedSelectedIndices[0]);
        }
      } else {
        addSiblingWithColumn(savedSelectedNode, savedSelectedIndices[0]);
      }
    } else {
      // Multiple column selection
      if (savedSelectionMode === 'child') {
        // 🔥 CRITICAL FIX: Create multiple children, one for each column
        savedSelectedIndices.forEach((columnIndex, i) => {
          const isLast = i === savedSelectedIndices.length - 1;
          addChildContainer(savedSelectedNode, !isLast); // Skip re-render except for last one
          
          // Find the newly created child and assign column - DO IT SYNCHRONOUSLY
          if (savedSelectedNode.children && savedSelectedNode.children.length > 0) {
            const newChild = savedSelectedNode.children[savedSelectedNode.children.length - 1];
            assignColumnToNode(newChild, columnIndex, !isLast); // Skip re-render except for last one
          }
        });
      } else {
        // Create multiple sibling containers
        savedSelectedIndices.forEach(columnIndex => {
          addSiblingWithColumn(savedSelectedNode, columnIndex);
        });
      }
    }
  });
}

// Helper function to get column name by index
function getColumnName(columnIndex) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  return sheetData?.headers?.[columnIndex] || `Column ${columnIndex + 1}`;
}

// Function to add sibling with column
function addSiblingWithColumn(node, columnIndex) {
  // Validate inputs
  if (!node) {
    console.error('❌ Node is null or undefined in addSiblingWithColumn');
    return;
  }
  
  if (columnIndex === undefined || columnIndex === null || isNaN(columnIndex)) {
    console.error('❌ Invalid columnIndex in addSiblingWithColumn:', columnIndex);
    return;
  }
  
  // First create a sibling using existing functionality
  addSiblingContainer(node);
  
  // Find the newly created sibling (it will be the last child of the parent)
  const parentNode = window.ExcelViewerNodeManager ? 
    window.ExcelViewerNodeManager.findParentNode(node.id) : 
    findParentNodeLocal(node);
    
  if (parentNode && parentNode.children && parentNode.children.length > 0) {
    const newSibling = parentNode.children[parentNode.children.length - 1];
    
    // Assign column to the new sibling
    setTimeout(() => {
      assignColumnToNode(newSibling, columnIndex);
    }, 100); // Small delay to ensure DOM is updated
  } else {
    console.error('Could not find newly created sibling');
  }
}

// Local fallback for finding parent node
function findParentNodeLocal(targetNode) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  if (!sheetData || !sheetData.root) return null;
  
  function searchInNode(node) {
    if (node.children) {
      for (let child of node.children) {
        if (child.id === targetNode.id) {
          return node;
        }
        const found = searchInNode(child);
        if (found) return found;
      }
    }
    return null;
  }
  
  return searchInNode(sheetData.root);
}

// Hide column selection area when clicking outside
document.addEventListener('click', (e) => {
  if (!currentSelectionArea) return;
  
  const addContentButtons = document.querySelectorAll('.add-content-button');
  
  // Check if click is on an add content button or inside the selection area
  let isInsideSelection = currentSelectionArea.contains(e.target);
  let isAddContentButton = Array.from(addContentButtons).some(btn => btn.contains(e.target));
  
  if (!isInsideSelection && !isAddContentButton) {
    hideColumnSelectionArea();
  }
}); 