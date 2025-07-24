/**
 * UI Controls Module - UI Controls, Styling & Initialization
 * Verantwoordelijk voor: Style controls, expand/collapse, initialization
 */

// Handle style controls
function applyStyles() {
  // Get DOM elements from core module
  const { fontSizeSelect, themeSelect, cellPaddingSelect, showGridCheckbox } = window.ExcelViewerCore.getDOMElements();
  
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
  // Get DOM elements from core module
  const { 
    tabButtons, 
    fontSizeSelect, 
    themeSelect, 
    cellPaddingSelect, 
    showGridCheckbox,
    exportPdfButton,
    previewPdfButton
  } = window.ExcelViewerCore.getDOMElements();
  
  // Get export and PDF functions from other modules
  const exportToPdf = window.ExcelViewerExport.exportToPdf;
  const previewPdf = window.ExcelViewerExport.previewPdf;
  const saveExportSelection = window.ExcelViewerCore.saveExportSelection;
  
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
  
  // Initialize toolbar including the Add Parent Level button (should be available from other modules)
  if (window.ExcelViewerStructureBuilder && window.ExcelViewerStructureBuilder.initToolbar) {
    window.ExcelViewerStructureBuilder.initToolbar();
  }
  
  // Apply initial styles
  applyStyles();
}

// Export functions for other modules
window.ExcelViewerUIControls = {
  applyStyles,
  toggleAllNodes,
  initialize
}; 