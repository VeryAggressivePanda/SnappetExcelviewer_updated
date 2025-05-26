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
  
  // Add smart column selector if node doesn't have data yet
  if (!node.columnIndex && node.columnIndex !== 0 && !node.isPlaceholder) {
    const columnSelect = document.createElement('select');
    columnSelect.className = 'column-selector';
    columnSelect.style.padding = '2px 6px';
    columnSelect.style.fontSize = '12px';
    columnSelect.style.border = '1px solid #ccc';
    columnSelect.style.borderRadius = '3px';
    
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = 'Select column...';
    columnSelect.appendChild(defaultOption);
    
    // Add available columns with smart filtering
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    if (sheetData && sheetData.headers) {
      const usedColumns = getUsedColumnsInHierarchy(node);
      
      sheetData.headers.forEach((header, index) => {
        if (!header) return;
        const option = document.createElement('option');
        option.value = index;
        
        if (usedColumns.includes(index)) {
          option.textContent = `${header} (already used)`;
          option.style.color = '#999';
          option.style.fontStyle = 'italic';
          option.disabled = true;
        } else {
          option.textContent = header;
        }
        
        columnSelect.appendChild(option);
      });
    }
    
    columnSelect.addEventListener('change', (e) => {
      e.stopPropagation();
      if (columnSelect.value) {
        console.log('Column selected in live-editor:', columnSelect.value, 'calling assignColumnToNode');
        // Directly assign the column and automatically determine values
        assignColumnToNode(node, parseInt(columnSelect.value));
      }
    });
    
    editControls.appendChild(columnSelect);
  }
  
  // Add child button
  const addChildBtn = document.createElement('button');
  addChildBtn.className = 'add-child-button';
  addChildBtn.textContent = '+ Child';
  addChildBtn.title = 'Add child container';
  addChildBtn.style.padding = '2px 6px';
  addChildBtn.style.fontSize = '12px';
  addChildBtn.style.border = '1px solid #4caf50';
  addChildBtn.style.borderRadius = '3px';
  addChildBtn.style.background = '#e8f5e9';
  addChildBtn.style.cursor = 'pointer';
  addChildBtn.style.color = '#4caf50';
  
  addChildBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    addChildContainer(node);
  });
  
  editControls.appendChild(addChildBtn);
  
  // Add sibling button (not for root)
  if (node.level > -1) {
    const addSiblingBtn = document.createElement('button');
    addSiblingBtn.className = 'add-sibling-button';
    addSiblingBtn.textContent = '+ Sibling';
    addSiblingBtn.title = 'Add sibling container';
    addSiblingBtn.style.padding = '2px 6px';
    addSiblingBtn.style.fontSize = '12px';
    addSiblingBtn.style.border = '1px solid #2196f3';
    addSiblingBtn.style.borderRadius = '3px';
    addSiblingBtn.style.background = '#e3f2fd';
    addSiblingBtn.style.cursor = 'pointer';
    addSiblingBtn.style.color = '#2196f3';
    
    addSiblingBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      addSiblingContainer(node);
    });
    
    editControls.appendChild(addSiblingBtn);
  }
  
  // Add delete button (not for root)
  if (node.level > -1 && !node.isPlaceholder) {
    const deleteBtn = document.createElement('button');
    deleteBtn.className = 'delete-button';
    deleteBtn.textContent = '×';
    deleteBtn.title = 'Delete container';
    deleteBtn.style.padding = '2px 8px';
    deleteBtn.style.fontSize = '16px';
    deleteBtn.style.border = '1px solid #f44336';
    deleteBtn.style.borderRadius = '3px';
    deleteBtn.style.background = '#ffebee';
    deleteBtn.style.cursor = 'pointer';
    deleteBtn.style.color = '#f44336';
    
    deleteBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      deleteContainer(node);
    });
    
    editControls.appendChild(deleteBtn);
  }
  
  // Add Layout Toggle button for all nodes (not just those with children)
  const layoutToggleBtn = document.createElement('button');
  layoutToggleBtn.className = 'layout-toggle-button';
  
  // Check current layout mode (default to horizontal)
  const isVertical = node.layoutMode === 'vertical';
  layoutToggleBtn.textContent = isVertical ? '⬇️' : '➡️';
  layoutToggleBtn.title = isVertical ? 'Switch to horizontal layout' : 'Switch to vertical layout';
  layoutToggleBtn.style.padding = '2px 6px';
  layoutToggleBtn.style.fontSize = '12px';
  layoutToggleBtn.style.border = '1px solid #ff9800';
  layoutToggleBtn.style.borderRadius = '3px';
  layoutToggleBtn.style.background = '#fff3e0';
  layoutToggleBtn.style.cursor = 'pointer';
  layoutToggleBtn.style.color = '#ff9800';
  
  layoutToggleBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    toggleNodeLayout(node);
  });
  
  editControls.appendChild(layoutToggleBtn);

  // Add Manage Items button for nodes with column data
  if ((node.columnIndex || node.columnIndex === 0) && !node.isPlaceholder) {
    const manageItemsBtn = document.createElement('button');
    manageItemsBtn.className = 'manage-items-button';
    manageItemsBtn.textContent = '☑';
    manageItemsBtn.title = 'Select values';
    manageItemsBtn.style.background = '#f3e5f5';
    manageItemsBtn.style.border = '1px solid #9c27b0';
    manageItemsBtn.style.color = '#9c27b0';
    manageItemsBtn.style.padding = '2px 6px';
    manageItemsBtn.style.fontSize = '12px';
    manageItemsBtn.style.borderRadius = '3px';
    manageItemsBtn.style.cursor = 'pointer';
    
    manageItemsBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      // Refresh/reprocess the column data automatically
      if (node.columnIndex !== undefined) {
        console.log('Refreshing column data for:', node.columnName);
        assignColumnToNode(node, node.columnIndex);
      }
    });
    
    editControls.appendChild(manageItemsBtn);
  }
  
  header.appendChild(editControls);
}

// Function to add a child container
function addChildContainer(parentNode) {
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
    layoutMode: 'horizontal' // Default to horizontal layout
  };
  
  if (!parentNode.children) {
    parentNode.children = [];
  }
  
  parentNode.children.push(newNode);
  
  // Check if this parent is the first sibling (template) and apply structure to all siblings
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  // Find the grandparent to check if parentNode is the first child
  const grandParent = findParentInStructure(sheetData.root, parentNode.id);
  if (grandParent && grandParent.children && grandParent.children.length > 1) {
    const parentIndex = grandParent.children.findIndex(child => child.id === parentNode.id);
    if (parentIndex === 0) {
      // This parent is the template (first sibling), apply structure to all siblings
      console.log('Parent is template, applying structure to all siblings...');
      applyTemplateStructureToAllSiblings(parentNode);
    }
  }
  
  // Re-render with proper styling
  if (window.renderSheet) {
    window.renderSheet(activeSheetId, sheetData);
  }
  
  // Apply flexbox styling after render
  applyFlexboxStyling();
}

// Function to add a sibling container
function addSiblingContainer(node) {
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
    layoutMode: 'horizontal' // Default to horizontal layout
  };
  
  if (!parent.children) {
    parent.children = [];
  }
  
  // Insert after current node
  const currentIndex = parent.children.findIndex(child => child.id === node.id);
  parent.children.splice(currentIndex + 1, 0, newNode);
  
  // Re-render with proper styling
  if (window.renderSheet) {
    window.renderSheet(activeSheetId, sheetData);
  }
  
  // Apply flexbox styling after render
  applyFlexboxStyling();
}

// Function to delete a container
function deleteContainer(node) {
  if (confirm(`Delete container "${node.value || node.columnName || 'New Container'}" and all its children?`)) {
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    
    // Find parent and remove node
    const parent = findParentInStructure(sheetData.root, node.id);
    if (parent && parent.children) {
      parent.children = parent.children.filter(child => child.id !== node.id);
    }
    
    // Re-render with proper styling
    if (window.renderSheet) {
      window.renderSheet(activeSheetId, sheetData);
    }
    
    // Apply flexbox styling after render
    applyFlexboxStyling();
  }
}

// Function to assign a column to a node (AUTOMATIC VALUE ASSIGNMENT!)
function assignColumnToNode(node, columnIndex) {
  console.log('assignColumnToNode called with:', { node: node.value, columnIndex });
  
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const headers = sheetData.headers;
  
  // Update node with column info
  node.columnIndex = columnIndex;
  node.columnName = headers[columnIndex] || `Column ${columnIndex + 1}`;
  
  // Get contextually filtered values
  const contextualValues = getContextualValuesForNode(node, columnIndex);
  
  console.log(`Found ${contextualValues.length} values for column ${node.columnName}:`, contextualValues);
  
  if (contextualValues.length === 0) {
    node.value = node.columnName;
    node.children = [];
  } else if (contextualValues.length === 1) {
    // Single value - assign it directly
    node.value = contextualValues[0];
    node.children = [];
  } else {
    // Multiple values - check if this is a root level node
    if (node.level === 0) {
      // Root level: Create siblings instead of children
      createSiblingsFromValues(node, contextualValues);
      return; // Exit early, don't call template function
    } else {
      // Child level: Create children for all contextual values
      node.value = node.columnName;
      node.children = contextualValues.map(value => ({
        id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        value: value,
        columnName: node.columnName,
        columnIndex: node.columnIndex,
        level: node.level + 1,
        children: [],
        properties: [],
        layoutMode: node.layoutMode || 'horizontal'
      }));
    }
  }
  
  // Apply template to siblings if this is the first sibling
  applyTemplateToSiblingsIfFirst(node);
  
  // Re-render with proper styling
  if (window.renderSheet) {
    window.renderSheet(activeSheetId, sheetData);
  }
  applyFlexboxStyling();
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

// Function to create siblings from column values (for root level nodes)
function createSiblingsFromValues(node, values) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  // Find the parent of this node
  const parent = findParentInStructure(sheetData.root, node.id);
  if (!parent || !parent.children) return;
  
  // Find the index of the current node
  const currentIndex = parent.children.findIndex(child => child.id === node.id);
  if (currentIndex === -1) return;
  
  // Update the current node with the first value
  node.value = values[0];
  node.children = [];
  
  // Create siblings for the remaining values
  const newSiblings = values.slice(1).map(value => ({
    id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
    value: value,
    columnName: node.columnName,
    columnIndex: node.columnIndex,
    level: node.level,
    children: [],
    properties: [],
    layoutMode: node.layoutMode || 'horizontal'
  }));
  
  // Insert the new siblings after the current node
  parent.children.splice(currentIndex + 1, 0, ...newSiblings);
  
  console.log(`Created ${newSiblings.length} siblings for values:`, values.slice(1));
  
  // Re-render with proper styling
  if (window.renderSheet) {
    window.renderSheet(activeSheetId, sheetData);
  }
  
  // Apply flexbox styling after render
  applyFlexboxStyling();
}

// Function to create multiple values as siblings
function createAsSiblings(node, selectedValues) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  // Find the parent of this node
  const parent = findParentInStructure(sheetData.root, node.id);
  if (!parent || !parent.children) return;
  
  // Find the index of the current node
  const currentIndex = parent.children.findIndex(child => child.id === node.id);
  if (currentIndex === -1) return;
  
  // Update the current node with the first value
  node.value = selectedValues[0];
  node.children = [];
  
  // Create siblings for the remaining values
  const newSiblings = selectedValues.slice(1).map(value => ({
    id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
    value: value,
    columnName: node.columnName,
    columnIndex: node.columnIndex,
    level: node.level,
    children: [],
    properties: [],
    layoutMode: node.layoutMode || 'horizontal'
  }));
  
  // Insert the new siblings after the current node
  parent.children.splice(currentIndex + 1, 0, ...newSiblings);
  
  console.log(`Created ${newSiblings.length} siblings for values:`, selectedValues.slice(1));
  
  // Apply template if this is the first sibling
  applyTemplateToSiblingsIfFirst(node);
  
  // Re-render with proper styling
  if (window.renderSheet) {
    window.renderSheet(activeSheetId, sheetData);
  }
  
  // Apply flexbox styling after render
  applyFlexboxStyling();
}

// Function to create multiple values as children
function createAsChildren(node, selectedValues) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  // Create child nodes for all selected values
  node.value = node.columnName;
  node.children = selectedValues.map(value => ({
    id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
    value: value,
    columnName: node.columnName,
    columnIndex: node.columnIndex,
    level: node.level + 1,
    children: [],
    properties: [],
    layoutMode: node.layoutMode || 'horizontal'
  }));
  
  console.log(`Created ${selectedValues.length} children for values:`, selectedValues);
  
  // Apply template if this is the first sibling
  applyTemplateToSiblingsIfFirst(node);
  
  // Re-render with proper styling
  if (window.renderSheet) {
    window.renderSheet(activeSheetId, sheetData);
  }
  
  // Apply flexbox styling after render
  applyFlexboxStyling();
}

// Function to show column values selector (redirects to inline selector)
function showColumnValuesSelector(node) {
  console.log('showColumnValuesSelector called - redirecting to inline selector');
  
  // Find the node element and its edit controls
  const nodeEl = document.querySelector(`[data-node-id="${node.id}"]`);
  if (nodeEl && node.columnIndex !== undefined) {
    const editControls = nodeEl.querySelector('.edit-controls');
    const manageBtn = editControls ? editControls.querySelector('.manage-items-button') : null;
    
    if (editControls && manageBtn) {
      showInlineValueSelector(node, node.columnIndex, editControls, manageBtn);
    } else {
      console.error('Could not find edit controls or manage button for inline selector');
    }
  } else {
    console.error('Could not find node element or node has no column index');
  }
}

// Function to apply selected column values
function applyColumnValues(node, selectedValues) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  if (selectedValues.length === 0) {
    node.value = node.columnName;
    node.children = [];
  } else if (selectedValues.length === 1) {
    // Single value - assign it directly
    node.value = selectedValues[0];
    node.children = [];
  } else {
    // Multiple values - create children for all values
    node.value = node.columnName;
    node.children = selectedValues.map(value => ({
      id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
      value: value,
      columnName: node.columnName,
      columnIndex: node.columnIndex,
      level: node.level + 1,
      children: [],
      properties: [],
      layoutMode: node.layoutMode || 'horizontal'
    }));
  }
  
  // Apply template to siblings if this is the first sibling
  applyTemplateToSiblingsIfFirst(node);
  
  // Re-render with proper styling
  if (window.renderSheet) {
    window.renderSheet(activeSheetId, sheetData);
  }
  applyFlexboxStyling();
}

// Universal function to apply template to siblings if this is the first sibling
function applyTemplateToSiblingsIfFirst(node) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  // Find the parent of this node
  const parent = findParentInStructure(sheetData.root, node.id);
  if (!parent || !parent.children) return;
  
  // Check if this is the first child (template)
  const nodeIndex = parent.children.findIndex(child => child.id === node.id);
  if (nodeIndex !== 0) return; // Only apply template if this is the first child
  
  console.log('Node is first sibling - applying template to all siblings...', node);
  
  // SIMPLIFIED: Only apply template structure to siblings at child level
  // Don't auto-generate anything at root level - user manually creates structure
  if (node.level > 0) {
    // Child level: Check if the parent has siblings, and if so, apply structure to all parent siblings
    const grandParent = findParentInStructure(sheetData.root, parent.id);
    if (grandParent && grandParent.children && grandParent.children.length > 1) {
      // The parent has siblings, so apply the template structure to all parent siblings
      const parentIndex = grandParent.children.findIndex(child => child.id === parent.id);
      if (parentIndex === 0) {
        // This parent is the template, apply structure to all siblings
        console.log('Parent is template, applying structure to all parent siblings...');
        applyTemplateStructureToAllSiblings(parent);
      }
    }
  }
}

// Function to apply template structure to all siblings (when first sibling is modified)
function applyTemplateStructureToAllSiblings(templateNode) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const rawData = getRawExcelData(activeSheetId);
  
  // Find the parent of this template node
  const parent = findParentInStructure(sheetData.root, templateNode.id);
  if (!parent || !parent.children) return;
  
  // Check if this is the first child (template)
  const templateIndex = parent.children.findIndex(child => child.id === templateNode.id);
  if (templateIndex !== 0) return; // Only apply template if this is the first child
  
  console.log('Applying template structure to ALL siblings...', templateNode);
  
  // Apply template structure to all other siblings
  for (let i = 1; i < parent.children.length; i++) {
    const siblingNode = parent.children[i];
    
    console.log(`Applying template structure to sibling ${i}: ${siblingNode.value}`);
    
    // Copy column assignment from template
    if (templateNode.columnIndex !== undefined && templateNode.columnIndex !== null) {
      siblingNode.columnIndex = templateNode.columnIndex;
      siblingNode.columnName = templateNode.columnName;
    }
    
    // Clone the template structure completely
    siblingNode.children = cloneTemplateStructureForSibling(templateNode.children, siblingNode);
    siblingNode.layoutMode = templateNode.layoutMode;
    
    // Apply column data to the cloned children
    populateClonedStructureWithData(siblingNode.children, siblingNode, rawData, 0);
  }
}

// Function to populate cloned structure with data for a specific sibling
function populateClonedStructureWithData(clonedChildren, parentSibling, rawData, depth = 0) {
  if (!clonedChildren || clonedChildren.length === 0) return;
  
  // Prevent infinite recursion - max depth of 5 levels
  if (depth > 5) {
    console.warn('Maximum recursion depth reached in populateClonedStructureWithData');
    return;
  }
  
  clonedChildren.forEach(clonedChild => {
    if (clonedChild.columnIndex !== undefined && clonedChild.columnIndex !== null) {
      // Populate this child with data specific to its parent sibling
      populateChildNodeWithSiblingData(clonedChild, parentSibling, rawData);
    }
    
    // Recursively populate any grandchildren that were already in the template
    // Don't recurse into newly created children to avoid infinite loops
    if (clonedChild.children && clonedChild.children.length > 0 && !clonedChild._newlyCreated) {
      populateClonedStructureWithData(clonedChild.children, clonedChild, rawData, depth + 1);
    }
  });
}

// Function to apply template structure to siblings (structure only, no data population)
function applyTemplateStructureToSiblings(templateNode) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const rawData = getRawExcelData(activeSheetId);
  
  // Find the parent of this template node
  const parent = findParentInStructure(sheetData.root, templateNode.id);
  if (!parent || !parent.children) return;
  
  // Check if this is the first child (template)
  const templateIndex = parent.children.findIndex(child => child.id === templateNode.id);
  if (templateIndex !== 0) return; // Only apply template if this is the first child
  
  console.log('Applying template structure to siblings with column data...', templateNode);
  
  // Apply template structure to all other siblings
  for (let i = 1; i < parent.children.length; i++) {
    const siblingNode = parent.children[i];
    
    console.log(`Applying template structure to sibling ${i}: ${siblingNode.value}`);
    
    // Clone the template structure
    siblingNode.children = cloneTemplateStructureForSibling(templateNode.children, siblingNode);
    siblingNode.layoutMode = templateNode.layoutMode;
    
    // Apply column data to the cloned children
    siblingNode.children.forEach(clonedChild => {
      if (clonedChild.columnIndex !== undefined && clonedChild.columnIndex !== null) {
        // Copy the column assignment from template
        clonedChild.columnName = templateNode.children.find(tc => tc.id !== clonedChild.id && tc.columnIndex === clonedChild.columnIndex)?.columnName || clonedChild.columnName;
        
        // Populate with data for this specific sibling
        populateChildNodeWithSiblingData(clonedChild, siblingNode, rawData);
      }
    });
  }
}

// Function to populate child node with data specific to its sibling parent
function populateChildNodeWithSiblingData(childNode, parentSibling, rawData) {
  if (!childNode.columnIndex && childNode.columnIndex !== 0) return;
  if (!parentSibling.columnIndex && parentSibling.columnIndex !== 0) return;
  
  console.log(`Populating child ${childNode.columnName} for parent ${parentSibling.value}`);
  
  // Find rows that match the parent sibling's value in its column
  const parentMatchingRows = rawData.filter((row, index) => {
    if (index === 0) return false; // Skip header row
    return row[parentSibling.columnIndex] && 
           row[parentSibling.columnIndex].toString().trim() === parentSibling.value;
  });
  
  if (parentMatchingRows.length === 0) {
    console.log(`No matching rows found for parent ${parentSibling.value}`);
    return;
  }
  
  // Get unique values for this child from the parent's matching rows
  const uniqueChildValues = new Set();
  parentMatchingRows.forEach(row => {
    const value = row[childNode.columnIndex];
    if (value && value.toString().trim()) {
      uniqueChildValues.add(value.toString().trim());
    }
  });
  
  const childValues = Array.from(uniqueChildValues).sort();
  console.log(`Found ${childValues.length} child values for ${childNode.columnName}:`, childValues);
  
  if (childValues.length === 0) {
    childNode.value = childNode.columnName;
    return;
  }
  
  // IMPORTANT: Check if the template child has a specific value selected
  // If so, we need to find the corresponding value for this parent sibling
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const templateParent = findTemplateParent(parentSibling, sheetData.root);
  
  if (templateParent && templateParent.children && templateParent.children.length > 0) {
    const templateChild = templateParent.children.find(child => 
      child.columnIndex === childNode.columnIndex && child.columnName === childNode.columnName
    );
    
    if (templateChild && templateChild.value && templateChild.value !== templateChild.columnName) {
      // Template child has a specific value selected
      // Check if this value exists for this parent sibling
      if (childValues.includes(templateChild.value)) {
        childNode.value = templateChild.value;
        console.log(`Applied template value ${templateChild.value} to child of ${parentSibling.value}`);
        return;
      } else {
        console.log(`Template value ${templateChild.value} not found for ${parentSibling.value}, using first available value`);
        if (childValues.length > 0) {
          childNode.value = childValues[0];
          return;
        }
      }
    }
  }
  
  if (childValues.length === 1) {
    // Single value - just update the child
    childNode.value = childValues[0];
  } else {
    // Multiple values - DON'T create child nodes automatically to prevent infinite recursion
    // Just use the first value
    childNode.value = childValues[0];
  }
}

// Function to automatically apply template structure to all siblings
function autoApplyTemplateToSiblings(templateNode) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const rawData = getRawExcelData(activeSheetId);
  
  if (!rawData || rawData.length <= 1) return;
  
  // Find the parent of this template node
  const parent = findParentInStructure(sheetData.root, templateNode.id);
  if (!parent || !parent.children) return;
  
  // Check if this is the first child (template) and has column data
  const templateIndex = parent.children.findIndex(child => child.id === templateNode.id);
  if (templateIndex !== 0) return; // Only apply template if this is the first child
  
  // Only proceed if the template node has column data
  if (templateNode.columnIndex === undefined || templateNode.columnIndex === null) return;
  
  console.log('Auto-applying template from first child to all siblings...', templateNode);
  
  // Get all unique values from the template's column
  const uniqueValues = new Set();
  for (let i = 1; i < rawData.length; i++) {
    const value = rawData[i][templateNode.columnIndex];
    if (value && value.toString().trim()) {
      uniqueValues.add(value.toString().trim());
    }
  }
  
  const allValues = Array.from(uniqueValues).sort();
  console.log(`Found ${allValues.length} unique values in column ${templateNode.columnIndex}:`, allValues);
  
  // Clear existing children except the template
  parent.children = [templateNode];
  
  // Create siblings for all other values
  allValues.forEach((value, index) => {
    if (index === 0) {
      // Update the template node with the first value
      templateNode.value = value;
      return;
    }
    
    // Create a new sibling for each additional value
    const siblingNode = {
      id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
      value: value,
      columnName: templateNode.columnName,
      columnIndex: templateNode.columnIndex,
      level: templateNode.level,
      children: cloneTemplateStructureForSibling(templateNode.children, templateNode),
      properties: [],
      layoutMode: templateNode.layoutMode || 'horizontal'
    };
    
    console.log(`Created sibling for value: ${value}`);
    parent.children.push(siblingNode);
    
    // Populate the cloned structure with data for this sibling
    populateNodeWithMatchingData(siblingNode, rawData);
  });
  
  // Also populate the template node with its data
  populateNodeWithMatchingData(templateNode, rawData);
}

// Function to clone template structure for a sibling
function cloneTemplateStructureForSibling(templateChildren, targetNode) {
  if (!templateChildren || templateChildren.length === 0) return [];
  
  return templateChildren.map(templateChild => {
    const clonedChild = {
      id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
      value: templateChild.value,
      columnName: templateChild.columnName,
      columnIndex: templateChild.columnIndex,
      level: templateChild.level,
      children: cloneTemplateStructureForSibling(templateChild.children, targetNode),
      properties: [],
      layoutMode: templateChild.layoutMode || 'horizontal'
    };
    
    console.log(`Cloned template child: ${templateChild.value} -> ${clonedChild.value}`);
    return clonedChild;
  });
}

// Function to populate node with matching data from Excel
function populateNodeWithMatchingData(node, rawData) {
  if (!node.columnIndex && node.columnIndex !== 0) return;
  
  // Find rows that match this node's value in its column
  const matchingRows = rawData.filter((row, index) => {
    if (index === 0) return false; // Skip header row
    return row[node.columnIndex] && row[node.columnIndex].toString().trim() === node.value;
  });
  
  console.log(`Found ${matchingRows.length} matching rows for ${node.value} (column ${node.columnIndex})`);
  
  // If this node has children, populate them with data from matching rows
  if (node.children && node.children.length > 0) {
    node.children.forEach(child => {
      console.log(`Populating child: ${child.value} with column ${child.columnIndex}`);
      populateChildWithData(child, matchingRows, rawData);
    });
  }
}

// Function to populate child nodes with data
function populateChildWithData(childNode, parentMatchingRows, allRawData) {
  if (!childNode.columnIndex && childNode.columnIndex !== 0) return;
  
  // Get unique values for this child from the parent's matching rows
  const uniqueChildValues = new Set();
  
  parentMatchingRows.forEach(row => {
    const value = row[childNode.columnIndex];
    if (value && value.toString().trim()) {
      uniqueChildValues.add(value.toString().trim());
    }
  });
  
  const childValues = Array.from(uniqueChildValues).sort();
  
  if (childValues.length === 0) return;
  
  if (childValues.length === 1) {
    // Single value - just update the child and add properties if it's a leaf node
    childNode.value = childValues[0];
    
    // If this child has no children, add properties from remaining columns
    if (!childNode.children || childNode.children.length === 0) {
      addPropertiesToLeafNode(childNode, parentMatchingRows, allRawData);
    } else {
      // If this child has children (template structure), populate them recursively
      childNode.children.forEach(grandchild => {
        const grandchildMatchingRows = parentMatchingRows.filter(row => 
          row[childNode.columnIndex] && 
          row[childNode.columnIndex].toString().trim() === childNode.value
        );
        populateChildWithData(grandchild, grandchildMatchingRows, allRawData);
      });
    }
  } else {
    // Multiple values - check if we have a template structure to clone
    if (childNode.children && childNode.children.length > 0) {
      // We have a template structure - clone it for each value
      const templateStructure = childNode.children;
      childNode.value = childNode.columnName;
      childNode.children = [];
      
      childValues.forEach(value => {
        // Clone the template structure for this value
        const clonedStructure = cloneTemplateStructureForSibling(templateStructure, childNode);
        
        // Create a container node for this value
        const valueNode = {
          id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
          value: value,
          columnName: childNode.columnName,
          columnIndex: childNode.columnIndex,
          level: childNode.level + 1,
          children: clonedStructure,
          properties: [],
          layoutMode: childNode.layoutMode || 'horizontal'
        };
        
        childNode.children.push(valueNode);
        
        // Populate the cloned structure with data for this specific value
        const valueMatchingRows = parentMatchingRows.filter(row => 
          row[childNode.columnIndex] && 
          row[childNode.columnIndex].toString().trim() === value
        );
        
        valueNode.children.forEach(clonedChild => {
          populateChildWithData(clonedChild, valueMatchingRows, allRawData);
        });
      });
    } else {
      // No template structure - create simple children
      childNode.value = childNode.columnName;
      childNode.children = childValues.map(value => ({
        id: `node-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        value: value,
        columnName: childNode.columnName,
        columnIndex: childNode.columnIndex,
        level: childNode.level + 1,
        children: [],
        properties: [],
        layoutMode: childNode.layoutMode || 'horizontal'
      }));
      
      // Recursively populate grandchildren
      childNode.children.forEach(grandchild => {
        const grandchildMatchingRows = parentMatchingRows.filter(row => 
          row[childNode.columnIndex] && 
          row[childNode.columnIndex].toString().trim() === grandchild.value
        );
        populateChildWithData(grandchild, grandchildMatchingRows, allRawData);
      });
    }
  }
}

// Function to add properties to leaf nodes from remaining columns
function addPropertiesToLeafNode(leafNode, matchingRows, allRawData) {
  if (matchingRows.length === 0) return;
  
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const headers = sheetData.headers;
  
  if (!headers) return;
  
  // Find the highest column index used in the structure
  const usedColumnIndices = findUsedColumnIndices(leafNode);
  
  // Add properties from columns that come after the leaf node's column
  const startColumn = Math.max(leafNode.columnIndex + 1, Math.max(...usedColumnIndices) + 1);
  
  leafNode.properties = [];
  
  for (let colIndex = startColumn; colIndex < headers.length; colIndex++) {
    if (!headers[colIndex]) continue;
    
    // Get values from this column for the matching rows
    const columnValues = matchingRows.map(row => row[colIndex] || '').filter(val => val.toString().trim());
    
    if (columnValues.length > 0) {
      leafNode.properties.push({
        column: colIndex,
        columnName: headers[colIndex],
        values: columnValues
      });
    }
  }
}

// Function to find all used column indices in the structure path to this node
function findUsedColumnIndices(node) {
  const indices = [];
  
  // Traverse up to find all parent column indices
  let currentNode = node;
  while (currentNode) {
    if (currentNode.columnIndex !== undefined && currentNode.columnIndex !== null) {
      indices.push(currentNode.columnIndex);
    }
    
    // Find parent (this is a simplified approach)
    const activeSheetId = window.excelData.activeSheetId;
    const sheetData = window.excelData.sheetsLoaded[activeSheetId];
    currentNode = findParentInStructure(sheetData.root, currentNode.id);
    
    if (currentNode && currentNode.type === 'root') break;
  }
  
  return indices;
}

// Helper function to find parent in structure
function findParentInStructure(node, targetId) {
  if (!node.children) return null;
  
  // Check if any direct child matches
  for (const child of node.children) {
    if (child.id === targetId) {
      return node;
    }
  }
  
  // Recursively check children
  for (const child of node.children) {
    const found = findParentInStructure(child, targetId);
    if (found) return found;
  }
  
  return null;
}

// Function to toggle layout mode for a node
function toggleNodeLayout(node) {
  // Toggle between horizontal and vertical layout
  node.layoutMode = node.layoutMode === 'vertical' ? 'horizontal' : 'vertical';
  
  console.log(`Toggled layout for node ${node.value || node.columnName} to ${node.layoutMode}`);
  
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  // Apply layout to ALL nodes with the same column name and level throughout the entire structure
  applyLayoutToAllSimilarNodes(node, sheetData.root);
  
  // Re-render to apply the new layout
  if (window.renderSheet) {
    window.renderSheet(activeSheetId, sheetData);
  }
  
  // Apply layout styling after render
  applyLayoutStyling();
}

// Function to apply layout to all similar nodes (same column name and level)
function applyLayoutToAllSimilarNodes(targetNode, rootNode) {
  const targetColumnName = targetNode.columnName;
  const targetLevel = targetNode.level;
  const targetLayoutMode = targetNode.layoutMode;
  
  console.log(`Applying layout ${targetLayoutMode} to all nodes with column "${targetColumnName}" at level ${targetLevel}`);
  
  function traverseAndApplyLayout(node) {
    // Check if this node matches our criteria
    if (node.columnName === targetColumnName && 
        node.level === targetLevel && 
        node.id !== targetNode.id) {
      
      console.log(`Found matching node: ${node.value} - applying layout ${targetLayoutMode}`);
      node.layoutMode = targetLayoutMode;
    }
    
    // Recursively check children
    if (node.children && node.children.length > 0) {
      node.children.forEach(child => traverseAndApplyLayout(child));
    }
  }
  
  // Start traversal from root
  traverseAndApplyLayout(rootNode);
}

// Function to apply layout styling based on node layout modes
function applyLayoutStyling() {
  console.log('Applying layout styling...');
  
  // Find all children containers and apply layout based on parent node's layoutMode
  const allContainers = document.querySelectorAll('[class*="-children"]');
  
  allContainers.forEach(container => {
    // Find the parent node element
    const parentNodeEl = container.closest('[data-node-id]');
    if (!parentNodeEl) return;
    
    const nodeId = parentNodeEl.getAttribute('data-node-id');
    if (!nodeId) return;
    
    // Find the actual node data
    const node = findNodeById(nodeId);
    if (!node) return;
    
    // Remove existing layout classes
    container.classList.remove(
      'layout-horizontal',
      'layout-vertical',
      'layout-single-child',
      'layout-two-children',
      'layout-three-children',
      'layout-many-children'
    );
    
    // Count children
    const childCount = container.children.length;
    
    // Apply layout based on node's layoutMode (default to horizontal) even if no children yet
    const layoutMode = node.layoutMode || 'horizontal';
    
    if (layoutMode === 'vertical') {
      container.classList.add('layout-vertical');
    } else {
      container.classList.add('layout-horizontal');
      
      // Add specific classes for horizontal layouts only if there are children
      if (childCount > 0) {
        if (childCount === 1) {
          container.classList.add('layout-single-child');
        } else if (childCount === 2) {
          container.classList.add('layout-two-children');
        } else if (childCount === 3) {
          container.classList.add('layout-three-children');
        } else {
          container.classList.add('layout-many-children');
        }
      }
    }
    
    console.log(`Applied ${layoutMode} layout to container with ${childCount} children`);
  });
}

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
  
  // Traverse up the hierarchy to find all used column indices
  let currentNode = node;
  while (currentNode && currentNode.level >= 0) {
    const parent = findParentInStructure(sheetData.root, currentNode.id);
    if (parent && (parent.columnIndex || parent.columnIndex === 0)) {
      usedColumns.push(parent.columnIndex);
    }
    currentNode = parent;
  }
  
  return usedColumns;
}

// Function to get contextual values for a node based on its parent hierarchy
function getContextualValuesForNode(node, columnIndex) {
  const activeSheetId = window.excelData.activeSheetId;
  const rawData = getRawExcelData(activeSheetId);
  
  console.log('getContextualValuesForNode called:', { 
    nodeValue: node.value, 
    columnIndex, 
    rawDataLength: rawData ? rawData.length : 0 
  });
  
  if (!rawData || rawData.length <= 1) {
    console.log('No raw data available');
    return [];
  }
  
  // Get the FULL parent context (all parents up the hierarchy)
  const fullParentContext = getFullParentContext(node);
  console.log('Full parent context:', fullParentContext);
  
  // Special handling for merged cells - find the range for each parent value
  let filteredRows = [];
  
  if (fullParentContext.length === 0) {
    // No parent context - get all values
    filteredRows = rawData.slice(1);
  } else {
    // Find rows that belong to the parent hierarchy, handling merged cells
    filteredRows = findRowsForParentHierarchy(rawData, fullParentContext);
  }
  
  console.log('Rows after filtering by parent hierarchy:', filteredRows.length);
  
  // Get unique values from the filtered rows for this column
  const uniqueValues = new Set();
  filteredRows.forEach(row => {
    const value = row[columnIndex];
    if (value && value.toString().trim()) {
      uniqueValues.add(value.toString().trim());
    }
  });
  
  const result = Array.from(uniqueValues).sort();
  console.log('Unique values found for column', columnIndex, ':', result);
  
  return result;
}

// Function to get parent context (column-value pairs from root to this node)
function getParentContext(node) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const context = [];
  
  // Traverse up the hierarchy to build context
  let currentNode = node;
  while (currentNode && currentNode.level >= 0) {
    const parent = findParentInStructure(sheetData.root, currentNode.id);
    if (parent && (parent.columnIndex || parent.columnIndex === 0) && parent.value && parent.value !== parent.columnName) {
      context.unshift({ // Add to beginning to maintain order from root
        columnIndex: parent.columnIndex,
        value: parent.value
      });
    }
    currentNode = parent;
  }
  
  return context;
}

// Function to get FULL parent context (all parents up the hierarchy)
function getFullParentContext(node) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const context = [];
  
  // Traverse up the hierarchy to build COMPLETE context
  let currentNode = node;
  while (currentNode && currentNode.level >= 0) {
    const parent = findParentInStructure(sheetData.root, currentNode.id);
    if (parent && 
        (parent.columnIndex !== undefined && parent.columnIndex !== null) && 
        parent.value && 
        parent.value !== parent.columnName &&
        parent.level >= 0) {
      context.unshift({ // Add to beginning to maintain order from root
        columnIndex: parent.columnIndex,
        value: parent.value
      });
    }
    currentNode = parent;
  }
  
  return context;
}

// Function to get immediate parent context (only direct parent)
function getImmediateParentContext(node) {
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  
  // Find the immediate parent
  const parent = findParentInStructure(sheetData.root, node.id);
  console.log('Found parent for node', node.value, ':', parent ? {
    value: parent.value,
    columnIndex: parent.columnIndex,
    columnName: parent.columnName,
    level: parent.level
  } : 'null');
  
  // Only use parent context if parent has actual data (not just column name) and a valid column
  if (parent && 
      (parent.columnIndex !== undefined && parent.columnIndex !== null) && 
      parent.value && 
      parent.value !== parent.columnName &&
      parent.level >= 0) {
    const result = {
      columnIndex: parent.columnIndex,
      value: parent.value
    };
    console.log('Returning parent context:', result);
    return result;
  }
  
  console.log('No valid parent context found - parent either has no column data or only shows column name');
  return null;
}

// Function to show inline value selector with checkboxes (NO MODAL!)
function showInlineValueSelector(node, columnIndex, editControls, oldSelect) {
  console.log('showInlineValueSelector called with:', { node, columnIndex, editControls, oldSelect });
  const activeSheetId = window.excelData.activeSheetId;
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  const headers = sheetData.headers;
  
  // Update node with column info
  node.columnIndex = columnIndex;
  node.columnName = headers[columnIndex] || `Column ${columnIndex + 1}`;
  
  // Get contextually filtered values
  const contextualValues = getContextualValuesForNode(node, columnIndex);
  
  if (contextualValues.length === 0) {
    node.value = node.columnName;
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
  
  // Fallback to sheet data
  const sheetData = window.excelData.sheetsLoaded[sheetId];
  if (sheetData && sheetData.rawData) {
    return sheetData.rawData;
  }
  
  return null;
}

// Function to find rows for parent hierarchy, handling merged cells correctly
function findRowsForParentHierarchy(rawData, parentContext) {
  if (!rawData || rawData.length <= 1 || parentContext.length === 0) {
    return rawData.slice(1);
  }
  
  const dataRows = rawData.slice(1); // Skip header
  const matchingRows = [];
  
  // Find the starting row that matches the full parent hierarchy
  let startRowIndex = -1;
  for (let i = 0; i < dataRows.length; i++) {
    const row = dataRows[i];
    let matchesAll = true;
    
    // Check if this row matches ALL parent context values
    for (const parentCtx of parentContext) {
      const cellValue = row[parentCtx.columnIndex];
      if (!cellValue || cellValue.toString().trim() !== parentCtx.value) {
        matchesAll = false;
        break;
      }
    }
    
    if (matchesAll) {
      startRowIndex = i;
      break;
    }
  }
  
  if (startRowIndex === -1) {
    console.log('No starting row found for parent hierarchy');
    return [];
  }
  
  console.log(`Found starting row at index ${startRowIndex} for hierarchy:`, parentContext);
  
  // Now find the range: from startRowIndex until we hit a new value in ANY parent column
  matchingRows.push(dataRows[startRowIndex]);
  
  for (let i = startRowIndex + 1; i < dataRows.length; i++) {
    const row = dataRows[i];
    let shouldStop = false;
    
    // Check if any parent column has a new non-empty value
    for (const parentCtx of parentContext) {
      const cellValue = row[parentCtx.columnIndex];
      if (cellValue && cellValue.toString().trim() && 
          cellValue.toString().trim() !== parentCtx.value) {
        // Found a new value in a parent column - stop here
        shouldStop = true;
        break;
      }
    }
    
    if (shouldStop) {
      console.log(`Stopping at row ${i} due to new parent value`);
      break;
    }
    
    // This row belongs to our range (either has matching values or empty cells due to merging)
    matchingRows.push(row);
  }
  
  console.log(`Found ${matchingRows.length} rows in range for parent hierarchy`);
  return matchingRows;
}

// Helper function to find the template parent (first sibling)
function findTemplateParent(currentParent, root) {
  // Find the parent of currentParent
  const grandParent = findParentInStructure(root, currentParent.id);
  if (!grandParent || !grandParent.children || grandParent.children.length === 0) return null;
  
  // Return the first child (template)
  return grandParent.children[0];
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
window.showColumnValuesSelector = showColumnValuesSelector;
window.showInlineValueSelector = showInlineValueSelector;
window.showSiblingOrChildrenChoice = showSiblingOrChildrenChoice;
window.createAsSiblings = createAsSiblings;
window.applyTemplateStructureToAllSiblings = applyTemplateStructureToAllSiblings;
window.applyTemplateToSiblingsIfFirst = applyTemplateToSiblingsIfFirst; 