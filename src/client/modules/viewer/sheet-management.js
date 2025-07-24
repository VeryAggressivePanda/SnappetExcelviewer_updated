/**
 * Sheet Management Module - Sheet Loading, Saving & Rendering
 * Verantwoordelijk voor: Sheet data loading, configuration saving, sheet rendering
 */

// Function to load sheet data with TEMPLATE SETUP
async function loadSheetData(sheetId) {
  try {
    console.log(`Loading data for sheet ${sheetId}`);
    
    // If already loaded, return the cached data
    if (window.excelData.sheetsLoaded[sheetId]) {
      renderSheet(sheetId, window.excelData.sheetsLoaded[sheetId]);
      if (window.ExcelViewerCore && window.ExcelViewerCore.populateExportColumnDropdown) {
        window.ExcelViewerCore.populateExportColumnDropdown();
      }
      return;
    }
    
    const response = await fetch(`/api/sheet/${window.excelData.fileId}/${sheetId}`);
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    
    const apiResponse = await response.json();
    console.log('Raw sheet data received:', apiResponse);
    
    // Store raw Excel data for later use (for add Les modal etc.)
    if (apiResponse.data && apiResponse.data.data && window.ExcelViewerStructureBuilder) {
      window.ExcelViewerStructureBuilder.storeRawExcelData(sheetId, apiResponse.data.data);
    }
    
    // Store the raw data directly for manual structure building
    window.excelData.sheetsLoaded[sheetId] = apiResponse.data;
    
    // Check if this is raw data that needs manual structure building
    const sheetData = window.excelData.sheetsLoaded[sheetId];
    if (sheetData.isRawData || sheetData.needsConfiguration) {
      console.log('Raw data detected - setting up for manual structure building');
      // Don't process raw data automatically - let user build manually
    } else if (sheetData && sheetData.root && sheetData.root.children && sheetData.root.children.length > 0) {
      console.log(`Found existing hierarchical data with ${sheetData.root.children.length} top-level nodes`);
      
      // Set up parent references and template markings for the loaded data
      if (window.ExcelViewerNodeManager) {
        console.log('🔄 Setting up parent references and template markings for loaded data...');
        window.ExcelViewerNodeManager.updateParentReferences(sheetData.root);
        window.ExcelViewerNodeManager.identifyAndMarkTemplates(sheetData.root);
      }
    }
    
    // Update the active sheet ID
    window.excelData.activeSheetId = sheetId;
    
    // Render the sheet
    renderSheet(sheetId, window.excelData.sheetsLoaded[sheetId]);
    
    // Always populate the export column dropdown after loading data
    if (window.ExcelViewerCore && window.ExcelViewerCore.populateExportColumnDropdown) {
      window.ExcelViewerCore.populateExportColumnDropdown();
    }
    
  } catch (error) {
    console.error('Error loading sheet data:', error);
    
    // Show error in the tab content
    const tabContent = document.getElementById(`sheet-${sheetId}`);
    if (tabContent) {
      tabContent.innerHTML = `
        <div class="loading-error">
          <h3>Error Loading Sheet</h3>
          <p>Could not load data for this sheet: ${error.message}</p>
          <button onclick="window.ExcelViewerSheetManager.loadSheetData('${sheetId}')" class="control-button">
            Try Again
          </button>
        </div>
      `;
    }
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

// Function to render a sheet with preserved state and TEMPLATE SETUP
function renderSheet(sheetId, sheetData) {
  if (!sheetData) {
    console.error('No sheet data provided for rendering');
    return;
  }
  
  // Initialize parent references and template markings for the hierarchy
  if (sheetData.root && window.ExcelViewerNodeManager) {
    console.log('🔄 Setting up parent references and template markings...');
    window.ExcelViewerNodeManager.updateParentReferences(sheetData.root);
    window.ExcelViewerNodeManager.identifyAndMarkTemplates(sheetData.root);
  }
  
  // Store current expand state before re-rendering
  const expandStateMap = window.ExcelViewerNodeManager ? 
    window.ExcelViewerNodeManager.storeExpandState() : {};
  
  // Find the correct sheet content container
  const tabContent = document.getElementById(`sheet-${sheetId}`);
  if (!tabContent) {
    console.error(`Tab content element not found for sheet ${sheetId}`);
    return;
  }
  
  // Hide the loading indicator if it exists
  const loadingIndicator = tabContent.querySelector('.loading-indicator');
  if (loadingIndicator) {
    loadingIndicator.style.display = 'none';
  }
  
  // Create or find the sheet container within the tab content
  let sheetContainer = tabContent.querySelector('.sheet-container');
  if (!sheetContainer) {
    // Create the sheet container if it doesn't exist
    sheetContainer = document.createElement('div');
    sheetContainer.className = 'sheet-container';
    tabContent.appendChild(sheetContainer);
    console.log(`Created sheet container for sheet ${sheetId}`);
  }
  
  // Clear the container
  sheetContainer.innerHTML = '';
  
  // Populate export column dropdown
  if (window.ExcelViewerCore && window.ExcelViewerCore.populateExportColumnDropdown) {
    window.ExcelViewerCore.populateExportColumnDropdown();
  }
  
  // Clear any error messages
  const errorDiv = tabContent.querySelector('.error-message');
  if (errorDiv) {
    errorDiv.remove();
  }
  
  try {
    // Use the root data structure directly (don't treat everything as raw data!)
    const currentData = sheetData.root;
    
    if (!currentData) {
      throw new Error('No root found in sheet data');
    }
    
    // Store global reference for hierarchical data
    window.hierarchicalData = currentData;
    
    // If we have children to render, render them
    if (currentData.children && currentData.children.length > 0) {
      console.log(`Rendering ${currentData.children.length} top-level nodes for sheet ${sheetId}`);
      
      currentData.children.forEach(childNode => {
        if (window.ExcelViewerNodeRenderer && window.ExcelViewerNodeRenderer.renderNode) {
          const renderedNode = window.ExcelViewerNodeRenderer.renderNode(childNode, childNode.level || 0, expandStateMap);
          if (renderedNode) {
            sheetContainer.appendChild(renderedNode);
          }
        }
      });
      
      // Enable live editing mode for all sheets
      sheetData.isEditable = true;
      
      // Enable live editing functionality
      if (window.enableLiveEditing) {
        window.enableLiveEditing(sheetId);
      }
      
    } else {
      // Only show "no data" if there really is no root or children
      console.log("No hierarchical data found, showing empty state");
      
      // Check if this is raw data that needs manual structure building
      if (sheetData.isRawData) {
        console.log("Raw data detected, setting up empty structure for building");
        
        // Create an empty structure for manual building
        const emptyStructure = {
          headers: sheetData.headers,
          root: {
            type: 'root',
            children: [],
            level: -1
          },
          isEditable: true
        };
        
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
        
        emptyStructure.root.children = [starterNode];
        
        // Update the sheet data
        window.excelData.sheetsLoaded[sheetId] = emptyStructure;
        
        // Render the starter node
        if (window.ExcelViewerNodeRenderer && window.ExcelViewerNodeRenderer.renderNode) {
          const renderedNode = window.ExcelViewerNodeRenderer.renderNode(starterNode, 0, expandStateMap);
          if (renderedNode) {
            sheetContainer.appendChild(renderedNode);
          }
        }
        
        // Enable live editing
        if (window.enableLiveEditing) {
          window.enableLiveEditing(sheetId);
        }
      } else {
        // Regular no data message
        const noDataDiv = document.createElement('div');
        noDataDiv.className = 'no-data-message';
        noDataDiv.textContent = 'No hierarchical data available for this sheet.';
        sheetContainer.appendChild(noDataDiv);
      }
    }
    
  } catch (error) {
    console.error('Error rendering sheet:', error);
    
    // Show error message to user
    const errorDiv = document.createElement('div');
    errorDiv.className = 'error-message';
    errorDiv.style.color = 'red';
    errorDiv.style.padding = '20px';
    errorDiv.style.border = '1px solid red';
    errorDiv.style.borderRadius = '4px';
    errorDiv.style.margin = '20px 0';
    errorDiv.innerHTML = `
      <h3>Error Loading Sheet Data</h3>
      <p>${error.message}</p>
      <p>Please try refreshing the page or uploading the Excel file again.</p>
    `;
    sheetContainer.appendChild(errorDiv);
  }
}

// Helper function to truncate text with middle ellipsis
function truncateMiddle(text, maxLength = 35) {
  if (!text || text.length <= maxLength) return text;
  
  const start = Math.floor(maxLength / 2);
  const end = Math.ceil(maxLength / 2);
  
  return text.substring(0, start) + '...' + text.substring(text.length - end);
}

// Export functions for other modules
window.ExcelViewerSheetManager = {
  loadSheetData,
  saveHierarchyConfiguration,
  saveCurrentHierarchyConfiguration,
  renderSheet,
  truncateMiddle
}; 