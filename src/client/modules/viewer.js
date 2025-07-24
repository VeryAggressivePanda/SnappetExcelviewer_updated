document.addEventListener('DOMContentLoaded', function() {
  // Initialize core module first
  window.ExcelViewerCore.initializeCore();
  
  // Get DOM elements from core module
  const {
    tabButtons,
    sheetContents,
    fontSizeSelect,
    themeSelect,
    cellPaddingSelect,
    showGridCheckbox,
    exportColumnSelect,
    previewPdfButton,
    exportPdfButton,
    collapseAllButton,
    expandAllButton
  } = window.ExcelViewerCore.getDOMElements();
  
  // Use core functions for export dropdown management
  const populateExportColumnDropdown = window.ExcelViewerCore.populateExportColumnDropdown;
  const saveExportSelection = window.ExcelViewerCore.saveExportSelection;
  
  // Use export module functions
  const exportToPdf = window.ExcelViewerExport.exportToPdf;
  const previewPdf = window.ExcelViewerExport.previewPdf;
  const showExactPdfPreview = window.ExcelViewerExport.showExactPdfPreview;
  
  // Use data-processing module functions
  const showHierarchyConfigModal = window.ExcelViewerDataProcessor.showHierarchyConfigModal;
  const processExcelData = window.ExcelViewerDataProcessor.processExcelData;
  const processChildrenWithCurrentValues = window.ExcelViewerDataProcessor.processChildrenWithCurrentValues;
  const cleanupNode = window.ExcelViewerDataProcessor.cleanupNode;
  const addPropertiesToNode = window.ExcelViewerDataProcessor.addPropertiesToNode;
  
  // Use sheet-management module functions
  const loadSheetData = window.ExcelViewerSheetManager.loadSheetData;
  const saveHierarchyConfiguration = window.ExcelViewerSheetManager.saveHierarchyConfiguration;
  const saveCurrentHierarchyConfiguration = window.ExcelViewerSheetManager.saveCurrentHierarchyConfiguration;
  const renderSheet = window.ExcelViewerSheetManager.renderSheet;
  const truncateMiddle = window.ExcelViewerSheetManager.truncateMiddle;
  
  // Use node-renderer module functions
  const renderNode = window.ExcelViewerNodeRenderer.renderNode;
  const showAddChildElementModal = window.ExcelViewerNodeRenderer.showAddChildElementModal;
  
  // Use ui-controls module functions
  const applyStyles = window.ExcelViewerUIControls.applyStyles;
  const toggleAllNodes = window.ExcelViewerUIControls.toggleAllNodes;
  const initialize = window.ExcelViewerUIControls.initialize;
  
  // Use node-management module functions
  const getExcelCoordinates = window.ExcelViewerNodeManager.getExcelCoordinates;
  const getExcelCellValue = window.ExcelViewerNodeManager.getExcelCellValue;
  const addElementDirectly = window.ExcelViewerNodeManager.addElementDirectly;
  const getSavedHierarchyConfiguration = window.ExcelViewerNodeManager.getSavedHierarchyConfiguration;
  const findParentNode = window.ExcelViewerNodeManager.findParentNode;
  const storeExpandState = window.ExcelViewerNodeManager.storeExpandState;
  const addNode = window.ExcelViewerNodeManager.addNode;
  const deleteNode = window.ExcelViewerNodeManager.deleteNode;
  
  // Use structure-builder module functions
  const showAddLesModal = window.ExcelViewerStructureBuilder.showAddLesModal;
  const storeRawExcelData = window.ExcelViewerStructureBuilder.storeRawExcelData;
  const getRawExcelData = window.ExcelViewerStructureBuilder.getRawExcelData;
  const addParentLevel = window.ExcelViewerStructureBuilder.addParentLevel;
  const ensureLevelStyles = window.ExcelViewerStructureBuilder.ensureLevelStyles;
  const updateLevels = window.ExcelViewerStructureBuilder.updateLevels;
  const addChildContainer = window.ExcelViewerStructureBuilder.addChildContainer;
  const addSiblingContainer = window.ExcelViewerStructureBuilder.addSiblingContainer;
  const deleteContainer = window.ExcelViewerStructureBuilder.deleteContainer;
  const initToolbar = window.ExcelViewerStructureBuilder.initToolbar;
  const showStructureBuilder = window.ExcelViewerStructureBuilder.showStructureBuilder;
  
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
  
  // Initialize the application
  initialize();
}); 