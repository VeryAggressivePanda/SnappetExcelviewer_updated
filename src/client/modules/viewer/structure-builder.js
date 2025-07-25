/**
 * Structure Builder Module - Advanced Hierarchy Building & Structure Management
 * Verantwoordelijk voor: Les modals, hierarchy building, structure builder, toolbars, advanced features
 */

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
    
    // Add the Les node to the Week node (should be available from node-management module)
    if (window.ExcelViewerNodeManager && window.ExcelViewerNodeManager.addNode) {
      window.ExcelViewerNodeManager.addNode(weekNode, newLesNode);
    }
    
    // Store the expand state before updating UI
    if (window.ExcelViewerNodeManager && window.ExcelViewerNodeManager.storeExpandState) {
      window.ExcelViewerNodeManager.storeExpandState();
    }
    
    // Close modal
    document.body.removeChild(modal);
    
    // Update the UI without refreshing the page
    const sheetContainer = document.getElementById('sheet-container');
    if (sheetContainer && window.ExcelViewerNodeRenderer && window.ExcelViewerNodeRenderer.renderNode) {
      sheetContainer.innerHTML = '';
      window.ExcelViewerNodeRenderer.renderNode(window.hierarchicalData, 0);
      if (window.ExcelViewerUIControls && window.ExcelViewerUIControls.applyStyles) {
        window.ExcelViewerUIControls.applyStyles();
      }
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
  if (window.ExcelViewerSheetManager && window.ExcelViewerSheetManager.renderSheet) {
    window.ExcelViewerSheetManager.renderSheet(activeSheetId, sheetData, undefined);
  }
  
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

// Function to add a child container (REDIRECTED TO LIVE EDITOR)
function addChildContainer(parentNode) {
  // Use the template-aware version from live-editor.js directly
  if (window.ExcelViewerLiveEditor && window.ExcelViewerLiveEditor.addChildContainer) {
    window.ExcelViewerLiveEditor.addChildContainer(parentNode);
  } else {
    console.error('Live editor addChildContainer not available');
  }
}

// Function to add a sibling container (REDIRECTED TO LIVE EDITOR)
function addSiblingContainer(node) {
  // Use the template-aware version from live-editor.js directly
  if (window.ExcelViewerLiveEditor && window.ExcelViewerLiveEditor.addSiblingContainer) {
    window.ExcelViewerLiveEditor.addSiblingContainer(node);
  } else {
    console.error('Live editor addSiblingContainer not available');
  }
}

// Function to delete a container (REDIRECTED TO LIVE EDITOR)
function deleteContainer(node) {
  // Use the template-aware version from live-editor.js directly
  if (window.ExcelViewerLiveEditor && window.ExcelViewerLiveEditor.deleteContainer) {
    window.ExcelViewerLiveEditor.deleteContainer(node);
  } else {
    console.error('Live editor deleteContainer not available');
  }
}

// Function to initialize toolbar
function initToolbar() {
  // Get DOM elements from core module
  const { collapseAllButton, expandAllButton } = window.ExcelViewerCore.getDOMElements();
  
  // Get toggleAllNodes from ui-controls module
  const toggleAllNodes = window.ExcelViewerUIControls.toggleAllNodes;
  
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
  if (window.ExcelViewerSheetManager && window.ExcelViewerSheetManager.renderSheet) {
    window.ExcelViewerSheetManager.renderSheet(activeSheetId, sheetData, undefined);
  }
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

// 🔧 SMART TEMPLATE REPLICATION: Copy structure but use contextual Excel data
function applyTemplateReplicationToAllDuplicates(templateNode) {
  console.log('🔧 SMART template replication (fallback) for:', templateNode.value);
  
  // Call the main GLOBAL replication function if available
  if (window.applyGlobalTemplateReplication) {
    window.applyGlobalTemplateReplication(templateNode);
  } else {
    console.log('⚠️ Main GLOBAL replication function not available');
  }
}

// Helper function to update children with contextual data
function updateChildrenWithContextualData(parentNode, children) {
  children.forEach(child => {
    child.parent = parentNode;
    
    // If this child has Excel data context, update its data based on parent's context
    if (child.excelCoordinates && parentNode.excelCoordinates) {
      // Update child's data based on parent's row context
      // This would involve re-fetching Excel data for this specific context
      // Implementation depends on specific Excel data structure
    }
    
    // Recursively update grandchildren
    if (child.children && child.children.length > 0) {
      updateChildrenWithContextualData(child, child.children);
    }
  });
}

// Initialize the structure builder
function initialize() {
  // Use initialize from ui-controls module
  if (window.ExcelViewerUIControls && window.ExcelViewerUIControls.initialize) {
    window.ExcelViewerUIControls.initialize();
  }
}

// Export functions for other modules
window.ExcelViewerStructureBuilder = {
  showAddLesModal,
  storeRawExcelData,
  getRawExcelData,
  findParentNode,
  addParentLevel,
  ensureLevelStyles,
  updateLevels,
  addChildContainer,
  addSiblingContainer,
  deleteContainer,
  initToolbar,
  showStructureBuilder,
  createStructureNode,
  extractStructure,
  applyCustomStructure,
  processStructureRow,
  applyTemplateReplicationToAllDuplicates,
  initialize
}; 