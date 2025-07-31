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
    const level = parseInt(node.getAttribute('data-level') || '0');
    const cssLevel = level < 0 ? `_${Math.abs(level)}` : level;
    const collapsedClass = `level-${cssLevel}-collapsed`;
    
    // Only collapse/expand nodes that have children
    const hasChildren = node.querySelector('[class*="-children"]');
    if (hasChildren) {
      if (collapse) {
        node.classList.add(collapsedClass);
      } else {
        node.classList.remove(collapsedClass);
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
    showEmptyCellsCheckbox,
    exportPdfButton,
    previewPdfButton
  } = window.ExcelViewerCore.getDOMElements();
  
  // Get template mode toggle buttons
  const templateModeToggle = document.getElementById('template-mode-toggle');
  
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
  showEmptyCellsCheckbox.addEventListener('change', toggleEmptyCells);
  
  // Initialize template mode toggle buttons
  if (templateModeToggle) {
    templateModeToggle.addEventListener('click', handleTemplateModeToggle);
    
    // Restore template mode from localStorage or default to list
    const savedMode = localStorage.getItem('templateMode') || 'list';
    setActiveToggleButton(templateModeToggle, savedMode);
    console.log(`🔄 Restored template mode from localStorage: ${savedMode}`);
  }
  
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
  
  // Apply initial empty cells visibility
  toggleEmptyCells();
}

// Function to toggle empty cells visibility
function toggleEmptyCells() {
  const { showEmptyCellsCheckbox } = window.ExcelViewerCore.getDOMElements();
  
  // Load saved preference if not triggered by user interaction
  if (arguments.length === 0) {
    const savedPreference = localStorage.getItem('showEmptyCells');
    if (savedPreference !== null) {
      showEmptyCellsCheckbox.checked = savedPreference === 'true';
    }
  }
  
  const showEmpty = showEmptyCellsCheckbox.checked;
  
  console.log('Toggling empty cells visibility:', showEmpty ? 'show' : 'hide');
  
  // Add or remove class from document body to control empty cell visibility
  if (showEmpty) {
    document.body.classList.remove('hide-empty-cells');
  } else {
    document.body.classList.add('hide-empty-cells');
  }
  
  // Save preference to localStorage
  localStorage.setItem('showEmptyCells', showEmpty.toString());
}

// Function to set active toggle button
function setActiveToggleButton(toggleContainer, value) {
  const buttons = toggleContainer.querySelectorAll('.toggle-btn');
  buttons.forEach(btn => {
    if (btn.dataset.value === value) {
      btn.classList.add('active');
    } else {
      btn.classList.remove('active');
    }
  });
}

// Function to handle template mode toggle clicks
function handleTemplateModeToggle(event) {
  if (!event.target.classList.contains('toggle-btn')) return;
  
  const selectedMode = event.target.dataset.value;
  const activeSheetId = window.excelData.activeSheetId;
  
  if (!activeSheetId) {
    alert('Please select a sheet first');
    return;
  }
  
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  if (!sheetData) {
    alert('No sheet data available');
    return;
  }
  
  // Update button states
  setActiveToggleButton(event.currentTarget, selectedMode);
  
  console.log(`🔄 Switching to ${selectedMode} mode`);
  
  if (selectedMode === 'list') {
    // Switch to list template mode
    renderListTemplate(sheetData);
  } else {
    // Switch back to hierarchy mode
    renderHierarchyMode(sheetData);
  }
  
  // Save preference
  localStorage.setItem('templateMode', selectedMode);
}

// Legacy function to handle template mode changes (for compatibility)
function handleTemplateModeChange() {
  const templateModeSelect = document.getElementById('template-mode');
  if (templateModeSelect) {
    // Legacy dropdown support
    const selectedMode = templateModeSelect.value;
    handleTemplateModeToggle({ 
      target: { dataset: { value: selectedMode }, classList: { contains: () => true } }, 
      currentTarget: document.getElementById('template-mode-toggle') 
    });
  }
}

// Function to render hierarchy mode (restore original structure)
function renderHierarchyMode(sheetData) {
  console.log('🔄 Switching to hierarchy mode...');
  
  // Restore original processed data if available
  if (sheetData.originalProcessedData) {
    // Handle both old format (direct root) and new format (object with root + headers)
    if (sheetData.originalProcessedData.root) {
      sheetData.headers = sheetData.originalProcessedData.headers;
      sheetData.root = sheetData.originalProcessedData.root;
    } else {
      // Legacy format - direct root
      sheetData.root = sheetData.originalProcessedData;
    }
    sheetData.isListTemplate = false;
    sheetData.processedData = null; // Clear list template data
    
    // Re-render the sheet
    const activeSheetId = window.excelData.activeSheetId;
    if (window.ExcelViewerSheetManager && window.ExcelViewerSheetManager.renderSheet) {
      window.ExcelViewerSheetManager.renderSheet(activeSheetId, sheetData);
    }
    
    // Update export dropdown for hierarchy mode
    if (window.ExcelViewerCore?.populateExportColumnDropdown) {
      window.ExcelViewerCore.populateExportColumnDropdown();
    }
    
    console.log('✅ Restored hierarchy mode');
  } else {
    console.warn('No original data to restore, creating fresh Start Building state');
    
    // Create fresh "Start Building" state
    sheetData.root = {
      type: 'root',
      children: [{
        id: `node-starter-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        value: 'Start Building',
        columnName: '',
        level: 0,
        children: [],
        properties: [],
        isPlaceholder: true
      }],
      level: -1
    };
    sheetData.isListTemplate = false;
    sheetData.isEditable = true;
    sheetData.processedData = null;
    
    // Re-render the sheet
    const activeSheetId = window.excelData.activeSheetId;
    if (window.ExcelViewerSheetManager && window.ExcelViewerSheetManager.renderSheet) {
      window.ExcelViewerSheetManager.renderSheet(activeSheetId, sheetData);
    }
    
    console.log('✅ Created fresh hierarchy mode with Start Building');
  }
}

// Export functions for other modules
window.ExcelViewerUIControls = {
  applyStyles,
  toggleAllNodes,
  toggleEmptyCells,
  initialize,
  renderListTemplate, // Add new function to exports
  handleTemplateModeChange,
  renderHierarchyMode,
  setupScrollSynchronization
};

// NEW: Render 3 Independent Level-0 Nodes (NO template/duplicate system)
function renderListTemplate(sheetData) {
    console.log('🏗️ Creating automatic list template with intelligent data processing...');
    
    // Store original data for potential restoration (including headers)
    if (!sheetData.originalProcessedData && sheetData.root) {
        sheetData.originalProcessedData = {
            root: JSON.parse(JSON.stringify(sheetData.root)),
            headers: sheetData.headers ? [...sheetData.headers] : null
        };
        console.log('💾 Stored original hierarchy data for restoration');
    }
    
    // Get raw Excel data for automatic processing
    let rawData = null;
    
    // Try multiple sources for raw data
    if (window.rawExcelData && window.excelData?.activeSheetId) {
        rawData = window.rawExcelData[window.excelData.activeSheetId];
        console.log('📊 Found raw data in window.rawExcelData:', rawData?.length);
    } else if (sheetData.rawData && Array.isArray(sheetData.rawData)) {
        rawData = sheetData.rawData;
        console.log('📊 Found raw data in sheetData.rawData:', rawData?.length);
    } else if (sheetData.data && Array.isArray(sheetData.data)) {
        rawData = sheetData.data;
        console.log('📊 Found raw data in sheetData.data:', rawData?.length);
    }
    
    if (!rawData || !Array.isArray(rawData) || rawData.length < 2) {
        console.error('❌ No raw Excel data available for automatic processing. Available sources:', {
            windowRawExcelData: !!window.rawExcelData,
            activeSheetId: window.excelData?.activeSheetId,
            sheetDataRawData: !!sheetData.rawData,
            sheetDataData: !!sheetData.data
        });
        return;
    }
    
    console.log('📊 Processing Excel data automatically...');
    const headers = rawData[0];
    const dataRows = rawData.slice(1);
    
    // Detect if there are multiple courses (check for Course column)
    console.log('🔍 Headers:', headers);
    const courseColIndex = headers.findIndex(h => 
        h && h.toLowerCase().includes('course')
    );
    console.log('📋 Course column index:', courseColIndex);
    
    let autoNodes = [];
    
    if (courseColIndex >= 0) {
        // Multiple courses detected - create sections for each course
        console.log('📚 Multiple courses detected - creating sections for each course');
        autoNodes = analyzeExcelDataWithCourses(dataRows, headers, courseColIndex);
        sheetData.hasCourses = true;
    } else {
        // Single course - use existing logic
        console.log('📚 Single course detected - using standard layout');
        const synchronizedRows = analyzeExcelData(dataRows, headers);
        console.log('🔍 Synchronized rows (already sorted by blok+week):', synchronizedRows);
        
        // Create columns in new order: Blok, Weken, Materialen
        autoNodes = [
            createSyncedColumn('Blok', 'blok', synchronizedRows),
            createSyncedColumn('Weken', 'weken', synchronizedRows),
            createSyncedColumn('Materialen', 'materials', synchronizedRows)
        ];
        sheetData.hasCourses = false;
    }
    
    // FORCE EXPAND ALL COLUMNS
    autoNodes.forEach(node => {
        node.collapsed = false;
        node.forceExpanded = true;
        console.log(`🔓 Forcing expand for column: ${node.value}`);
    });
    
    // Replace current sheet data with automatic nodes
    sheetData.root = {
        type: 'root',
        children: autoNodes,
        level: -1
    };
    sheetData.processedData = autoNodes;
    sheetData.isListTemplate = true;
    
    // Re-render the sheet with new data
    const activeSheetId = window.excelData?.activeSheetId;
    if (activeSheetId && window.ExcelViewerSheetManager?.renderSheet) {
        window.ExcelViewerSheetManager.renderSheet(activeSheetId, sheetData);
        
        // Add scroll synchronization after render
        setTimeout(() => {
            setupScrollSynchronization();
        }, 100);
        
        // Update export dropdown for list template
        if (window.ExcelViewerCore?.populateExportColumnDropdown) {
            window.ExcelViewerCore.populateExportColumnDropdown();
        }
    }
    
    console.log('✅ Created automatic list template with 3 intelligent columns');
}

// Setup scroll synchronization between the 3 columns
function setupScrollSynchronization() {
    const listContainer = document.querySelector('.sheet-container[data-is-list-template="true"]');
    if (!listContainer) {
        console.log('❌ No list template container found for scroll sync');
        return;
    }
    
    const columnContents = listContainer.querySelectorAll('.level-0-content');
    if (columnContents.length !== 3) {
        console.log('❌ Expected 3 columns for scroll sync, found:', columnContents.length);
        return;
    }
    
    console.log('🔄 Setting up scroll synchronization for 3 columns');
    
    // Remove any existing event listeners
    columnContents.forEach(content => {
        const newContent = content.cloneNode(true);
        content.parentNode.replaceChild(newContent, content);
    });
    
    // Get fresh references after cloning
    const syncedColumns = listContainer.querySelectorAll('.level-0-content');
    let isScrolling = false;
    
    syncedColumns.forEach((column, index) => {
        column.addEventListener('scroll', function(e) {
            if (isScrolling) return;
            
            isScrolling = true;
            const scrollTop = this.scrollTop;
            
            // Sync other columns
            syncedColumns.forEach((otherColumn, otherIndex) => {
                if (otherIndex !== index) {
                    otherColumn.scrollTop = scrollTop;
                }
            });
            
            // Reset flag after a short delay
            setTimeout(() => {
                isScrolling = false;
            }, 10);
        });
    });
    
    console.log('✅ Scroll synchronization active for 3 columns');
}

// Apply curriculum abbreviations
function applyCurriculumAbbreviations(text) {
    // Return original text without abbreviations (as requested)
    return text || '';
}

// Analyze Excel data with multiple courses - create sections for each course
function analyzeExcelDataWithCourses(dataRows, headers, courseColIndex) {
    const autoNodes = [];
    
    // Group rows by course with forward fill logic
    const courseGroups = new Map();
    let lastKnownCourse = '';
    
    dataRows.forEach((row, index) => {
        if (!row || row.length === 0) return;
        
        // Forward fill course name like we do for Blok/Week/Les
        if (row[courseColIndex] && row[courseColIndex].toString().trim()) {
            lastKnownCourse = row[courseColIndex].toString().trim();
        }
        
        const courseName = lastKnownCourse || 'Onbekende Course';
        
        if (!courseGroups.has(courseName)) {
            courseGroups.set(courseName, []);
        }
        courseGroups.get(courseName).push(row);
    });
    
    console.log(`📚 Found ${courseGroups.size} courses:`, Array.from(courseGroups.keys()));
    
    // For each course, create the same 3-column structure as single course
    courseGroups.forEach((courseRows, courseName) => {
        console.log(`🔄 Processing course: ${courseName} with ${courseRows.length} rows`);
        
        // Analyze data for this specific course
        const synchronizedRows = analyzeExcelDataForCourse(courseRows, headers, courseColIndex);
        
        // Store synchronized rows globally for pagination calculations
        window.currentCourseSynchronizedRows = synchronizedRows;
        
        if (synchronizedRows.length > 0) {
            // Create course container with new column order: Blok, Weken, Materialen  
            const courseContainer = {
                id: `course-container-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
                value: courseName,
                columnName: '',
                columnIndex: '',
                level: -1,  // Course containers op level -1
                children: [
                    createSyncedColumn('Blok', 'blok', synchronizedRows),
                    createSyncedColumn('Weken', 'weken', synchronizedRows),
                    createSyncedColumn('Materialen', 'materials', synchronizedRows)
                ],
                isCourseContainer: true,
                collapsed: false,
                forceExpanded: true
            };
            
            autoNodes.push(courseContainer);
        }
    });
    
    return autoNodes;
}

// Analyze Excel data to create synchronized row-based data for 3 columns (single course version)
function analyzeExcelDataForCourse(dataRows, headers, courseColIndex = -1) {
    // Adjust column indexes if course column exists
    const blokColIndex = courseColIndex >= 0 ? 1 : 0;
    const weekColIndex = courseColIndex >= 0 ? 2 : 1;
    const lesColIndex = courseColIndex >= 0 ? 3 : 2;
    
    return analyzeExcelDataInternal(dataRows, headers, blokColIndex, weekColIndex, lesColIndex);
}

// Analyze Excel data to create synchronized row-based data for 3 columns  
function analyzeExcelData(dataRows, headers) {
    return analyzeExcelDataInternal(dataRows, headers, 0, 1, 2);
}

// Internal function that does the actual analysis
function analyzeExcelDataInternal(dataRows, headers, blokColIndex, weekColIndex, lesColIndex) {
    const materialGroupings = new Map(); // Group by material and blok
    
    // Find both material columns - Klas and Leerlingen  
    const klasColIndex = headers.findIndex(h => 
        h && h.toLowerCase().includes('instructie') && h.toLowerCase().includes('klas')
    );
    const leerlingenColIndex = headers.findIndex(h => 
        h && h.toLowerCase().includes('instructie') && h.toLowerCase().includes('leerlingen')
    );
    
    console.log(`📚 Found Klas column at index: ${klasColIndex}`);
    console.log(`📚 Found Leerlingen column at index: ${leerlingenColIndex}`);
    
    // Forward fill state - remember last known values for empty cells
    let lastKnownBlok = 'Onbekend';
    let lastKnownWeek = 'Onbekend';
    
    // First pass: collect all material-blok-week combinations
    dataRows.forEach((row, index) => {
        if (!row || row.length === 0) return;
        
        // Extract blok and week info with forward fill logic
        if (row[blokColIndex] && row[blokColIndex].toString().trim()) {
            lastKnownBlok = row[blokColIndex].toString().trim();
        }
        if (row[weekColIndex] && row[weekColIndex].toString().trim()) {
            lastKnownWeek = row[weekColIndex].toString().trim();
        }
        
        // Use full text without abbreviations (as requested)
        const blokFull = lastKnownBlok || '';
        const weekFull = lastKnownWeek || '';
        
        // Extract and combine material text from both Klas and Leerlingen columns
        const klasText = klasColIndex >= 0 ? (row[klasColIndex] || '') : '';
        const leerlingenText = leerlingenColIndex >= 0 ? (row[leerlingenColIndex] || '') : '';
        
        // Combine both texts with comma separation if both exist
        let combinedText = '';
        if (klasText.trim() && leerlingenText.trim()) {
            combinedText = `${klasText.trim()}, ${leerlingenText.trim()}`;
        } else if (klasText.trim()) {
            combinedText = klasText.trim();
        } else if (leerlingenText.trim()) {
            combinedText = leerlingenText.trim();
        }
        
        if (combinedText.trim()) {
            // Extract materials - handle Klas (klassikaal) and Leerlingen (per groepje) differently
            let materialsWithQuantity = [];
            
            // Process Klas text (klassikaal - no colon splitting)
            if (klasText.trim()) {
                const klasMaterials = extractMaterialsSimple(klasText.trim());
                materialsWithQuantity.push(...klasMaterials);
            }
            
            // Process Leerlingen text (per groepje - with colon splitting)
            if (leerlingenText.trim()) {
                const leerlingenMaterials = extractMaterialsWithQuantity(leerlingenText.trim(), '');
                materialsWithQuantity.push(...leerlingenMaterials);
            }
            
            // Group by material name and blok
            materialsWithQuantity.forEach(({materialName, quantity, fullMaterialText}) => {
                const materialKey = `${materialName}|${blokFull}`; // material + blok combination
                
                if (!materialGroupings.has(materialKey)) {
                    materialGroupings.set(materialKey, {
                        materialName: materialName,
                        blok: blokFull,
                        weken: [weekFull],
                        rowIndex: index,
                        originalText: fullMaterialText
                    });
                } else {
                    // Add week if not already present
                    const existing = materialGroupings.get(materialKey);
                    if (!existing.weken.includes(weekFull)) {
                        existing.weken.push(weekFull);
                    }
                }
            });
        }
    });
    
    // Second pass: create hierarchical structure grouped by blok+week
    const synchronizedRows = [];
    
    // Reorganize data: group by blok first, then by week within each blok
    const blokWeekGroupings = new Map(); // Key: "blok|week", Value: materials array
    
    materialGroupings.forEach((entry) => {
        entry.weken.forEach(week => {
            const blokWeekKey = `${entry.blok}|${week}`;
            
            if (!blokWeekGroupings.has(blokWeekKey)) {
                blokWeekGroupings.set(blokWeekKey, {
                    blok: entry.blok,
                    week: week,
                    materials: []
                });
            }
            
            blokWeekGroupings.get(blokWeekKey).materials.push(entry.materialName);
        });
    });
    
    // Sort blok+week combinations for proper display order
    const sortedBlokWeekEntries = Array.from(blokWeekGroupings.entries()).sort((a, b) => {
        const [blokWeekA] = a;
        const [blokWeekB] = b;
        const [blokA, weekA] = blokWeekA.split('|');
        const [blokB, weekB] = blokWeekB.split('|');
        
        // Extract numeric parts for proper sorting
        const getBlokNumber = (blok) => {
            const match = blok.match(/\d+/);
            return match ? parseInt(match[0]) : 0;
        };
        
        const getWeekNumber = (week) => {
            const match = week.match(/\d+/);
            return match ? parseInt(match[0]) : 0;
        };
        
        const blokNumA = getBlokNumber(blokA);
        const blokNumB = getBlokNumber(blokB);
        
        // First sort by blok number
        if (blokNumA !== blokNumB) {
            return blokNumA - blokNumB;
        }
        
        // Then sort by week number within same blok
        return getWeekNumber(weekA) - getWeekNumber(weekB);
    });
    
    // Create synchronized rows with hierarchical blok+week structure - apply firstForBlok AFTER sorting
    let currentBlok = '';
    sortedBlokWeekEntries.forEach(([blokWeekKey, data]) => {
        const { blok, week, materials } = data;
        
        // Combine all materials for this blok+week combination
        const uniqueMaterials = [...new Set(materials)]; // Remove duplicates
        const materialsDisplay = uniqueMaterials.join(', ');
        
        // Add to synchronized rows with full blok name (will fix this after sorting)
        synchronizedRows.push({
            material: materialsDisplay,
            blok: blok, // Keep full blok name for now
            weken: week,
            rowIndex: 0,
            isFirstForBlok: false // Will set this correctly after
        });
    });
    
    // Keep all blok names in the data - hiding will be handled per page after rendering
    synchronizedRows.forEach((row, index) => {
        row.isFirstForBlok = false; // Will be determined per page later
    });
    
    console.log(`📊 Created ${synchronizedRows.length} synchronized rows grouped by blok+week`);
    
    return synchronizedRows;
}

// Smart split function that respects parentheses
function smartSplit(text, separator) {
    const parts = [];
    let currentPart = '';
    let parenDepth = 0;
    
    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        
        if (char === '(') {
            parenDepth++;
            currentPart += char;
        } else if (char === ')') {
            parenDepth--;
            currentPart += char;
        } else if (separator.test(char) && parenDepth === 0) {
            // Split here - we're not inside parentheses
            parts.push(currentPart);
            currentPart = '';
        } else {
            currentPart += char;
        }
    }
    
    // Add the last part
    if (currentPart) {
        parts.push(currentPart);
    }
    
    return parts;
}

// Extract materials from klassikale text (no colon splitting, just comma/semicolon separated)
function extractMaterialsSimple(text) {
    const materials = [];
    
    // Smart split that respects parentheses - don't split on commas inside ()
    const parts = smartSplit(text, /[,;]/).map(part => part.trim()).filter(part => part.length > 0);
    
    parts.forEach(part => {
        // Extract quantity and material name: "10 fiches" -> quantity=10, material="fiches"
        const quantityMatch = part.match(/^(\d+)\s+(.+)$/);
        
        if (quantityMatch) {
            // Found quantity + material
            const quantity = quantityMatch[1];
            const materialName = quantityMatch[2]
                .trim()
                .replace(/\s+/g, ' '); // Normalize spaces
            
            if (materialName.length > 1) { // Lower threshold for klassikale materials
                materials.push({
                    materialName: materialName,
                    quantity: quantity,
                    fullMaterialText: part
                });
            }
        } else {
            // No quantity found, just material name
            const materialName = part
                .trim()
                .replace(/^\d+\s*/, '') // Remove any leading numbers
                .replace(/\s+/g, ' '); // Normalize spaces
            
            if (materialName.length > 1) { // Lower threshold for klassikale materials
                materials.push({
                    materialName: materialName,
                    quantity: '', // No quantity
                    fullMaterialText: part
                });
            }
        }
    });
    
    return materials.length > 0 ? materials : [];
}

// Extract material names with quantities from instruction text
function extractMaterialsWithQuantity(text, groupForm) {
    const materials = [];
    
    // Remove group form pattern from text to focus on materials
    let cleanText = text.replace(new RegExp(groupForm, 'gi'), '');
    
    // Split ONLY by commas and semicolons to keep material names intact
    // Don't split on colons or hyphens to preserve "vierkante vouwblaadjes"
    // Also clean up text after colon to get the actual materials list
    if (cleanText.includes(':')) {
        cleanText = cleanText.split(':').slice(1).join(':').trim(); // Take everything after the first colon
    }
    
    // No abbreviations here - they belong in curriculum processing
    
    const parts = smartSplit(cleanText, /[,;]/).map(part => part.trim()).filter(part => part.length > 0);
    
    parts.forEach(part => {
        // Extract quantity and material name: "10 fiches" -> quantity=10, material="fiches"
        const quantityMatch = part.match(/^(\d+)\s+(.+)$/);
        
        if (quantityMatch) {
            // Found quantity + material
            const quantity = quantityMatch[1];
            const materialName = quantityMatch[2]
                .trim()
                .replace(/\s+/g, ' '); // Normalize spaces
            
            if (materialName.length > 2 && !/^\d+$/.test(materialName)) {
                materials.push({
                    materialName: materialName,
                    quantity: quantity,
                    fullMaterialText: part
                });
            }
        } else {
            // No quantity found, just material name
            const materialName = part
                .trim()
                .replace(/^\d+\s*/, '') // Remove any leading numbers
                .replace(/\s+/g, ' '); // Normalize spaces
            
            if (materialName.length > 2 && !/^\d+$/.test(materialName)) {
                materials.push({
                    materialName: materialName,
                    quantity: '', // No quantity
                    fullMaterialText: part
                });
            }
        }
    });
    
    return materials.length > 0 ? materials : [{
        materialName: 'materiaal onbekend',
        quantity: '',
        fullMaterialText: 'materiaal onbekend'
    }];
}

// Create synchronized column node
function createSyncedColumn(title, type, synchronizedRows) {
    console.log(`🏗️ Creating ${title} column (type: ${type}) with ${synchronizedRows.length} rows`);
    
    const node = {
        id: `synced-${type}-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        value: title,
        columnName: '',
        columnIndex: '',
        level: 0,
        children: [],
        isListColumn: true,
        columnType: 'synchronized',
        autoType: type,
        collapsed: false,
        forceExpanded: true
    };
    
    // Create synchronized children - each row corresponds across columns
    synchronizedRows.forEach((rowData, index) => {
        let displayValue = '';
        
        switch (type) {
            case 'materials':
                displayValue = rowData.material; // This is already empty for non-first rows
                break;
            case 'blok':
                displayValue = rowData.blok;
                break;
            case 'weken':
                displayValue = rowData.weken;
                break;
            // Legacy support
            case 'groupForms':
                displayValue = rowData.groupForm;
                break;
            case 'curriculum':
                displayValue = rowData.curriculum;
                break;
        }
        
        // For empty materials, use a visible character that won't be filtered
        if (type === 'materials' && displayValue === '') {
            displayValue = '•'; // Bullet point for empty material rows
        }
        
        node.children.push({
            id: `synced-${type}-row-${index}-${Date.now()}`,
            value: displayValue,
            level: 1,
            children: [],
            syncRowIndex: index,  // Keep track of row synchronization
            isEmpty: displayValue === '', // Only empty string is truly empty, not bullet points
            isBullet: displayValue === '•' // Mark nodes that contain only bullet points
        });
    });
    
    console.log(`📝 Created synchronized ${type} column with ${node.children.length} rows`);
    return node;
} 