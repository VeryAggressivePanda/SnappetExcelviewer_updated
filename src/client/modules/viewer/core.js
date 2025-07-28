/**
 * Core Module - DOM Setup & Basic State Management
 * Verantwoordelijk voor: DOM elements, global data store, export dropdown
 */

// DOM Elements - alleen declareren, vullen in initialize
let tabButtons, sheetContents, fontSizeSelect, themeSelect, cellPaddingSelect;
let showGridCheckbox, showEmptyCellsCheckbox, exportColumnSelect, previewPdfButton, exportPdfButton;
let collapseAllButton, expandAllButton;

// Global data store - preserve existing properties but ensure all required properties exist
function initializeGlobalState() {
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
}

// Initialize DOM elements
function initializeDOMElements() {
  tabButtons = document.querySelectorAll('.tab-button');
  sheetContents = document.querySelectorAll('.sheet-content');
  fontSizeSelect = document.getElementById('font-size');
  themeSelect = document.getElementById('theme');
  cellPaddingSelect = document.getElementById('cell-padding');
  showGridCheckbox = document.getElementById('show-grid');
  showEmptyCellsCheckbox = document.getElementById('show-empty-cells');
  exportColumnSelect = document.getElementById('export-column');
  previewPdfButton = document.getElementById('preview-pdf');
  exportPdfButton = document.getElementById('export-pdf');
  collapseAllButton = document.getElementById('collapse-all');
  expandAllButton = document.getElementById('expand-all');
}

// Function to populate export column dropdown
function populateExportColumnDropdown() {
  const activeSheetId = window.excelData.activeSheetId;
  
  // Clear existing options
  exportColumnSelect.innerHTML = '';
  
  // Disable and clear if no active sheet or sheets not loaded
  if (!activeSheetId || !window.excelData.sheetsLoaded[activeSheetId]) {
    exportColumnSelect.disabled = true;
    previewPdfButton.disabled = true;
    exportPdfButton.disabled = true;
    return;
  }
  
  // Enable dropdown if active sheet is valid
  exportColumnSelect.disabled = false;
  previewPdfButton.disabled = false;
  exportPdfButton.disabled = false;
  
  // Get node values from active sheet data
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  if (!sheetData || !sheetData.root || !sheetData.root.children) {
    return;
  }
  
  // Get children of containers for the dropdown (not the containers themselves)
  let exportableNodes = [];
  
  // Loop through container nodes and get their children
  sheetData.root.children.forEach(containerNode => {
    if (containerNode.children && containerNode.children.length > 0) {
      // Get content children (not placeholders)
      const contentChildren = containerNode.children.filter(child => 
        !child.isPlaceholder && 
        child.value && 
        child.value.trim() !== ''
      );
      exportableNodes.push(...contentChildren);
    }
  });
  
  // If no children found, fall back to direct nodes
  if (exportableNodes.length === 0) {
    exportableNodes = sheetData.root.children.filter(node => 
      !node.isPlaceholder && 
      node.value && 
      node.value.trim() !== ''
    );
  }
  
  console.log("Exportable nodes (children of containers):", exportableNodes.map(n => ({ value: n.value, level: n.level })));
  
  // Add default option
  const defaultOption = document.createElement('option');
  defaultOption.value = '';
  defaultOption.text = 'Select an item to export...';
  exportColumnSelect.appendChild(defaultOption);
  
  // Add option for each exportable node  
  exportableNodes.forEach(node => {
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
      previewPdfButton.disabled = false;
      exportPdfButton.disabled = false;
    }
  }
}

// Function to save export selection to localStorage
function saveExportSelection(sheetId, value) {
  if (sheetId && value) {
    localStorage.setItem(`exportSelection_${sheetId}`, value);
  }
}

// Initialize core functionality
function initializeCore() {
  initializeGlobalState();
  initializeDOMElements();
}

// Export functions for other modules
window.ExcelViewerCore = {
  initializeCore,
  populateExportColumnDropdown,
  saveExportSelection,
  // Expose DOM elements for other modules
  getDOMElements: () => ({
    tabButtons,
    sheetContents,
    fontSizeSelect,
    themeSelect,
    cellPaddingSelect,
    showGridCheckbox,
    showEmptyCellsCheckbox,
    exportColumnSelect,
    previewPdfButton,
    exportPdfButton,
    collapseAllButton,
    expandAllButton
  })
}; 