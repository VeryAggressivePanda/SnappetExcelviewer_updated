/**
 * Sheet Management Module - Sheet Loading, Saving & Rendering
 * Verantwoordelijk voor: Sheet data loading, configuration saving, sheet rendering
 */

// Auto-detect hierarchy from Excel headers
function detectAutomaticHierarchy(headers) {
  console.log('🔍 Detecting automatic hierarchy from headers:', headers);
  
  if (!headers || !Array.isArray(headers)) {
    console.log('❌ No valid headers found for auto-detection');
    return null;
  }
  
  const hierarchy = {};
  
  // Look for standard column patterns
  const blokIndex = headers.findIndex(h => h && h.toLowerCase().includes('blok'));
  const weekIndex = headers.findIndex(h => h && h.toLowerCase().includes('week'));
  const lesIndex = headers.findIndex(h => h && h.toLowerCase().includes('les'));
  
  console.log('🔍 Found column indices:', { blokIndex, weekIndex, lesIndex });
  
  if (blokIndex >= 0 && weekIndex >= 0 && lesIndex >= 0) {
    // Standard hierarchy: Blok → Week → Les
    hierarchy[blokIndex] = null; // Top level
    hierarchy[weekIndex] = blokIndex; // Week is child of Blok
    hierarchy[lesIndex] = weekIndex; // Les is child of Week
    
    console.log('✅ Auto-detected standard hierarchy: Blok → Week → Les');
    return hierarchy;
  } else if (weekIndex >= 0 && lesIndex >= 0) {
    // Simple hierarchy: Week → Les
    hierarchy[weekIndex] = null; // Top level
    hierarchy[lesIndex] = weekIndex; // Les is child of Week
    
    console.log('✅ Auto-detected simple hierarchy: Week → Les');
    return hierarchy;
  } else {
    console.log('❌ Could not auto-detect standard hierarchy');
    return null;
  }
}

// Function to load sheet data with TEMPLATE SETUP
async function loadSheetData(sheetId) {
  try {
    console.log(`Loading data for sheet ${sheetId}`);
    
    // If already loaded, return the cached data
    if (window.excelData.sheetsLoaded[sheetId]) {
      renderSheet(sheetId, window.excelData.sheetsLoaded[sheetId], undefined);
      if (window.ExcelViewerCore && window.ExcelViewerCore.populateExportColumnDropdown) {
        window.ExcelViewerCore.populateExportColumnDropdown();
      }
      // Sync template mode dropdown with actual sheet mode
      syncTemplateModeDropdown();
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
    
    // Check if this is raw data - just set up empty structure for manual column selection
    const sheetData = window.excelData.sheetsLoaded[sheetId];
    
    // Only set default mode for completely fresh sheets that have no mode set yet
    if (sheetData.isListTemplate === undefined) {
      const currentMode = localStorage.getItem('templateMode') || 'list';
      console.log(`Setting initial mode for fresh sheet: ${currentMode}`);
      sheetData.isListTemplate = (currentMode === 'list');
      
      // If list mode is default, trigger list template rendering
      if (currentMode === 'list' && window.ExcelViewerUIControls?.renderListTemplate) {
        console.log('🎯 Triggering list template rendering for fresh sheet');
        window.ExcelViewerUIControls.renderListTemplate(sheetData);
        return; // Skip hierarchy setup
      }
    }
    
    // Set up hierarchy structure only if not in list template mode
    if (!sheetData.isListTemplate) {
      console.log('Setting up empty hierarchy structure for manual column selection');
      
      // Create an empty structure for manual building via existing column selectors
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
      sheetData.isRawData = false;
      sheetData.isEditable = true;
    }
    
    // Update the active sheet ID
    window.excelData.activeSheetId = sheetId;
    
    // Render the sheet
    renderSheet(sheetId, window.excelData.sheetsLoaded[sheetId], undefined);
    
    // Enable live editing AFTER rendering (so DOM exists)
    const currentSheetData = window.excelData.sheetsLoaded[sheetId];
    if (currentSheetData && currentSheetData.isEditable && window.enableLiveEditing) {
      setTimeout(() => {
        window.enableLiveEditing(sheetId);
        console.log('🎯 Live editing enabled for hierarchy sheet');
      }, 100);
    }
    
    // Always populate the export column dropdown after loading data
    if (window.ExcelViewerCore && window.ExcelViewerCore.populateExportColumnDropdown) {
      window.ExcelViewerCore.populateExportColumnDropdown();
    }
    
    // Sync template mode dropdown with actual sheet mode
    syncTemplateModeDropdown();
    
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

// Function to synchronize template mode toggle buttons with actual sheet state
function syncTemplateModeDropdown() {
  const templateModeToggle = document.getElementById('template-mode-toggle');
  if (!templateModeToggle) return;
  
  const activeSheetId = window.excelData.activeSheetId;
  if (!activeSheetId) return;
  
  const sheetData = window.excelData.sheetsLoaded[activeSheetId];
  if (!sheetData) return;
  
  // Determine actual mode based on sheet data
  const actualMode = sheetData.isListTemplate ? 'list' : 'hierarchy';
  
  // Get currently active button
  const activeButton = templateModeToggle.querySelector('.toggle-btn.active');
  const currentMode = activeButton ? activeButton.dataset.value : null;
  
  // Only update if different to avoid unnecessary updates
  if (currentMode !== actualMode) {
    console.log(`🔄 Syncing toggle from "${currentMode}" to actual mode "${actualMode}"`);
    
    // Update button states
    const buttons = templateModeToggle.querySelectorAll('.toggle-btn');
    buttons.forEach(btn => {
      if (btn.dataset.value === actualMode) {
        btn.classList.add('active');
      } else {
        btn.classList.remove('active');
      }
    });
  }
}

// Function to render a sheet with preserved state and TEMPLATE SETUP
function renderSheet(sheetId, sheetData) {
    console.log(`🔄 Loading data for sheet ${sheetId}`);
    
    // Find or create sheet container
    let sheetContainer = document.querySelector(`.sheet-content#sheet-${sheetId} .sheet-container`);
    if (!sheetContainer) {
        const sheetContentDiv = document.getElementById(`sheet-${sheetId}`);
        if (!sheetContentDiv) {
            console.error(`No sheet div found for ID: sheet-${sheetId}`);
            return;
        }

        // Hide loading indicator
        const loadingIndicator = sheetContentDiv.querySelector('.loading-indicator');
        if (loadingIndicator) {
            loadingIndicator.style.display = 'none';
        }

        // Create new sheet container
        sheetContainer = document.createElement('div');
        sheetContainer.className = 'sheet-container';
        sheetContentDiv.appendChild(sheetContainer);
        console.log(`🏗️ Created sheet container for sheet ${sheetId}`);
    }

    // Store sheet data globally for access by other components
    if (!window.sheetData) {
        window.sheetData = {};
    }
    window.sheetData[sheetId] = sheetData;

    // Apply list template styling if needed
    if (sheetData.isListTemplate) {
        sheetContainer.setAttribute('data-is-list-template', 'true');
        console.log('🏗️ Applied list template styling to sheet container');
        
        // Add course indicator if needed
        if (sheetData.hasCourses) {
            sheetContainer.setAttribute('data-has-courses', 'true');
            console.log('📚 Applied course styling to sheet container');
        } else {
            sheetContainer.removeAttribute('data-has-courses');
        }
    } else {
        sheetContainer.removeAttribute('data-is-list-template');
        sheetContainer.removeAttribute('data-has-courses');
    }

    // Clear existing content
    sheetContainer.innerHTML = '';

    // Render nodes
    let nodesToRender = [];
    
    if (sheetData.isListTemplate && sheetData.processedData) {
        // NEW: Automatic list template with processedData
        nodesToRender = sheetData.processedData;
        console.log(`🎯 Rendering automatic list template with ${nodesToRender.length} columns`);
    } else if (sheetData.root && sheetData.root.children) {
        // Original hierarchy mode
        nodesToRender = sheetData.root.children;
        console.log(`📊 Rendering hierarchy with ${nodesToRender.length} top-level nodes`);
    } else {
        console.log('📋 No data to render for sheet', sheetId);
        return;
    }

    // Set up parent references and template markings before rendering
    console.log('🔄 Setting up parent references and GLOBAL template markings...');
    setupParentReferences(nodesToRender);
    
    // Save expanded states before rendering
    if (window.ExcelViewerNodeManager && window.ExcelViewerNodeManager.saveExpandedStates) {
        window.ExcelViewerNodeManager.saveExpandedStates();
    }

    // Store exportable nodes reference before rendering for global access
    function collectExportableNodes(nodes) {
        const exportable = [];
        nodes.forEach(node => {
            if (node.level === 0) {  // Only include container nodes
                exportable.push(node);
            }
        });
        return exportable;
    }

    window.exportableNodes = collectExportableNodes(nodesToRender);
    console.log('🎯 Exportable nodes (children of containers):', window.exportableNodes);

    console.log(`🔧 Rendering ${nodesToRender.length} top-level nodes for sheet ${sheetId}`);

    // For list templates, create A4 pages with repeating headers
    if (sheetData.isListTemplate && nodesToRender.length > 0) {
        console.log('📄 Creating A4 pages for list template');
        console.log('Total nodes to render:', nodesToRender.length);
        
        // Check if we have course containers (level -1) or direct columns (level 0)
        const hasCourses = nodesToRender.some(node => node.level === -1 && node.isCourseContainer);
        console.log('Has course containers:', hasCourses);
        
        if (hasCourses) {
            // Course selection will be handled by populateExportColumnDropdown in core.js
            
            // Handle course containers - create A4 pages for each course
            nodesToRender.forEach((courseContainer, courseIndex) => {
                if (!courseContainer.isCourseContainer || !courseContainer.children) return;
                
                console.log(`📚 Processing course: ${courseContainer.value}`);
                const courseColumns = courseContainer.children; // The 3 columns within this course
                
                // Determine how many rows this course has
                const firstColumn = courseColumns[0];
                const totalRows = firstColumn && firstColumn.children ? firstColumn.children.length : 0;
                
                if (totalRows === 0) {
                    console.log(`No rows found for course ${courseContainer.value} - skipping`);
                    return;
                }
                
                // NIEUWE AANPAK: Use same proven logic as single course
                console.log('🚀 MULTI COURSE: Using same proven logic as single course for:', courseContainer.value);
                
                // Use actual measured height ratio: test measures 887px but real uses 527px  
                // Compensate: 900 * (887/527) = 1515px to account for measurement difference
                const compensatedHeight = Math.round(900 * (887/527));
                // Reduce height for title on first page of each course (approx 80px for "Materialenlijst" + sheet name)
                const heightWithTitle = compensatedHeight - 80;
                const pageBreaks = calculateRealPageBreaks(courseColumns, heightWithTitle, false); // false = multi course
                console.log('🚀 MULTI COURSE: Calculated page breaks:', pageBreaks);
                
                // Nu de echte pagina's maken met deze page breaks - DIRECT, geen timeouts
                const totalPages = pageBreaks.length;
                console.log(`Course ${courseContainer.value}: Splitting ${totalRows} rows across ${totalPages} A4 pages`);
                
                        // Create A4 pages for this course
        for (let pageIndex = 0; pageIndex < totalPages; pageIndex++) {
            const pageDiv = document.createElement('div');
            pageDiv.className = 'a4-page';
            pageDiv.setAttribute('data-page', `course-${courseIndex}-page-${pageIndex + 1}`);
            pageDiv.setAttribute('data-course', courseContainer.value);
            
            // Get sheet name
            const sheetTab = document.querySelector(`.tab-button[data-sheet-id="${sheetId}"]`);
            const sheetName = sheetTab ? sheetTab.textContent.trim() : 'Materialen';
            
            // Create main title (first page of EACH course)
            if (pageIndex === 0) {
                const mainTitle = document.createElement('div');
                mainTitle.textContent = 'Materialenlijst';
                mainTitle.style.cssText = `
                    color: #34a3d7;
                    font-size: 2rem;
                    font-weight: bold;
                    font-family: 'ABeZeh-Bold', sans-serif;
                    margin: 0;
                    line-height: 1.2;
                `;
                pageDiv.appendChild(mainTitle);
                
                const subTitle = document.createElement('div');
                subTitle.textContent = sheetName;
                subTitle.style.cssText = `
                    color: #34a3d7;
                    font-size: 1.5rem;
                    font-weight: bold;
                    font-family: 'ABeZeh-Bold', sans-serif;
                    margin: 0 0 1rem 0;
                    line-height: 1.2;
                `;
                pageDiv.appendChild(subTitle);
            }
            
            // Create course header
            const courseHeader = document.createElement('div');
            courseHeader.className = 'course-page-header';
            // Extract group part from different course name patterns:
            // "Snappet Rekenen (Groep 4)" -> "Groep 4"
            // "02. RekenWereld groep 3" -> "groep 3"
            // "05. Instruct! 2 groep 3" -> "groep 3"
            let groupText = courseContainer.value;
            const parenthesesMatch = courseContainer.value.match(/\(([^)]+)\)/);
            if (parenthesesMatch) {
                groupText = parenthesesMatch[1];
            } else {
                const groepMatch = courseContainer.value.match(/groep\s+\d+/i);
                if (groepMatch) {
                    groupText = groepMatch[0];
                }
            }
            courseHeader.textContent = groupText;
            courseHeader.style.cssText = `
                background: #34a3d7;
                color: white;
                padding: 8px 16px;
                border-radius: 50px;
                font-size: 1rem;
                font-weight: bold;
                width: fit-content;
                margin-bottom: 0.5rem;
                flex-shrink: 0;
            `;
            pageDiv.appendChild(courseHeader);
                    
                    // Create columns container
                    const columnsContainer = document.createElement('div');
                    columnsContainer.style.cssText = 'display: flex; flex-direction: row; gap: 0; flex: 1; min-height: 0;';
                    
                    // Use real page breaks instead of calculated rows
                    const startRowIndex = pageIndex === 0 ? 0 : pageBreaks[pageIndex - 1];
                    const endRowIndex = pageBreaks[pageIndex];
                    
                    // Create filtered columns for this page
                    courseColumns.forEach(column => {
                        const columnForPage = {
                            ...column,
                            children: column.children ? column.children.slice(startRowIndex, endRowIndex) : [],
                            id: `${column.id}-course-${courseIndex}-page-${pageIndex + 1}`
                        };
                        
                        if (window.ExcelViewerNodeRenderer && window.ExcelViewerNodeRenderer.renderNode) {
                            const columnElement = window.ExcelViewerNodeRenderer.renderNode(columnForPage, columnsContainer);
                            if (columnElement) {
                                columnsContainer.appendChild(columnElement);
                            }
                        }
                    });
                    
                    pageDiv.appendChild(columnsContainer);
                    sheetContainer.appendChild(pageDiv);
                    
                    console.log(`✅ Created page ${pageIndex + 1}/${totalPages} for course ${courseContainer.value}`);
                }
             });
            
            // Synchronize row heights for course-based layout after rendering is complete
            setTimeout(() => {
                synchronizeListTemplateRowHeights(sheetContainer);
                hideRepeatingBlokNamesPerPage(sheetContainer);
            }, 100);
            return; // Exit early for course handling
        }
        
        // Original logic for non-course (single course) templates
        // Export dropdown will be handled by populateExportColumnDropdown in core.js
        
        const firstColumn = nodesToRender[0];
        const totalRows = firstColumn && firstColumn.children ? firstColumn.children.length : 0;
        
        console.log('Total rows found in first column:', totalRows);
        
        if (totalRows === 0) {
            console.error('No rows found in columns - falling back to normal rendering');
            // Fall back to normal rendering
            nodesToRender.forEach((node) => {
                if (window.ExcelViewerNodeRenderer && window.ExcelViewerNodeRenderer.renderNode) {
                    const nodeElement = window.ExcelViewerNodeRenderer.renderNode(node, sheetContainer);
                    if (nodeElement) {
                        sheetContainer.appendChild(nodeElement);
                    }
                }
            });
            return;
        }
        
        console.log('🚀 SINGLE COURSE - Using real row height measurement');
        console.log('🚀 SINGLE COURSE: nodesToRender length:', nodesToRender.length);
        
        // NIEUWE AANPAK: Render alle rijen en meet echte hoogtes
        const pageBreaks = calculateRealPageBreaks(nodesToRender, 750, true); // true = single course, reduced height for more content
        console.log('🚀 SINGLE COURSE: Calculated page breaks:', pageBreaks);
        
        // Create separate A4 pages for each page break
        const totalPages = pageBreaks.length;
        for (let pageIndex = 0; pageIndex < totalPages; pageIndex++) {
            console.log(`🎯 Starting creation of page ${pageIndex + 1}/${totalPages}`);
            
            // Create A4 page container with CSS class
            const pageDiv = document.createElement('div');
            pageDiv.className = 'a4-page';
            pageDiv.setAttribute('data-page', pageIndex + 1);
            
            // Get sheet name
            const sheetTab = document.querySelector(`.tab-button[data-sheet-id="${sheetId}"]`);
            const sheetName = sheetTab ? sheetTab.textContent.trim() : 'Materialen';
            
            // Create main title (only on first page)
            if (pageIndex === 0) {
                const titleContainer = document.createElement('div');
                titleContainer.style.cssText = `
                    width: 100%;
                    margin-bottom: 1rem;
                    flex-shrink: 0;
                `;
                
                const mainTitle = document.createElement('div');
                mainTitle.textContent = 'Materialenlijst';
                mainTitle.style.cssText = `
                    color: #34a3d7;
                    font-size: 2rem;
                    font-weight: bold;
                    font-family: 'ABeZeh-Bold', sans-serif;
                    margin: 0;
                    line-height: 1.2;
                `;
                
                const subTitle = document.createElement('div');
                subTitle.textContent = sheetName;
                subTitle.style.cssText = `
                    color: #34a3d7;
                    font-size: 1.5rem;
                    font-weight: bold;
                    font-family: 'ABeZeh-Bold', sans-serif;
                    margin: 0 0 0 0;
                    line-height: 1.2;
                `;
                
                titleContainer.appendChild(mainTitle);
                titleContainer.appendChild(subTitle);
                pageDiv.appendChild(titleContainer);
            }
            
            // Create columns wrapper for single course with 1 1 3 flex ratio
            const columnsWrapper = document.createElement('div');
            columnsWrapper.style.cssText = 'display: flex; flex-direction: row; flex: 1; gap: 0; width: 100%;';
            pageDiv.appendChild(columnsWrapper);
            
            console.log(`Created pageDiv for page ${pageIndex + 1}`, pageDiv);
            
            // Use real page breaks instead of calculated rows
            const startRowIndex = pageIndex === 0 ? 0 : pageBreaks[pageIndex - 1];
            const endRowIndex = pageBreaks[pageIndex];
            
            console.log(`Page ${pageIndex + 1}: Will include rows ${startRowIndex} to ${endRowIndex - 1}`);
            
            // For each column, create a filtered version with only the rows for this page
            const columnsForPage = nodesToRender.map(column => {
                // Create header row for this column
                const columnHeaderNode = {
                    ...column,
                    children: [], // Clear children - we'll add filtered ones
                    id: `${column.id}-page-${pageIndex + 1}-header`,
                    isPageHeader: true
                };
                
                // Add only the rows that belong to this page
                if (column.children) {
                    const rowsForThisPage = column.children.slice(startRowIndex, endRowIndex);
                    columnHeaderNode.children = rowsForThisPage.map(row => ({
                        ...row,
                        id: `${row.id}-page-${pageIndex + 1}` // Make unique for this page
                    }));
                }
                
                return columnHeaderNode;
            });
            
            console.log(`Page ${pageIndex + 1}: Created ${columnsForPage.length} columns with ${endRowIndex - startRowIndex} rows each`);
            
            // Render all columns for this page
            columnsForPage.forEach((columnNode, columnIndex) => {
                if (window.ExcelViewerNodeRenderer && window.ExcelViewerNodeRenderer.renderNode) {
                    const columnElement = window.ExcelViewerNodeRenderer.renderNode(columnNode, columnsWrapper);
                    if (columnElement) {
                        columnElement.style.position = 'relative';
                        columnElement.style.marginBottom = '10px';
                        columnsWrapper.appendChild(columnElement);
                    } else {
                        console.warn(`Failed to render column ${columnIndex} on page ${pageIndex + 1}`);
                    }
                } else {
                    console.error('NodeRenderer not available for columns');
                }
            });
            
            // Add page to sheet container
            sheetContainer.appendChild(pageDiv);
            console.log(`✅ Page ${pageIndex + 1} completed and added to sheet container with ${columnsForPage.length} columns and ${endRowIndex - startRowIndex} rows`);
        }
        
        console.log(`✅ FINISHED: Created ${totalPages} A4 pages total using real row height measurements`);
        console.log('Sheet container now has', sheetContainer.children.length, 'children');
        
        // Synchronize row heights across columns for consistent zebra stripes after rendering is complete
        setTimeout(() => {
            synchronizeListTemplateRowHeights(sheetContainer);
            hideRepeatingBlokNamesPerPage(sheetContainer);
        }, 100);
        
    } else {
        // Normal rendering for non-list templates
        nodesToRender.forEach((node, index) => {
            if (window.ExcelViewerNodeRenderer && window.ExcelViewerNodeRenderer.renderNode) {
                const nodeElement = window.ExcelViewerNodeRenderer.renderNode(node, sheetContainer);
                if (nodeElement) {
                    sheetContainer.appendChild(nodeElement);
                }
            } else {
                console.error('NodeRenderer not available');
            }
        });
    }

    // Restore expanded states after rendering
    if (window.ExcelViewerNodeManager && window.ExcelViewerNodeManager.restoreExpandedStates) {
        window.ExcelViewerNodeManager.restoreExpandedStates();
    }
}

// Setup parent references for nodes (used in both hierarchy and automatic modes)
function setupParentReferences(nodes, parent = null) {
    if (!Array.isArray(nodes)) return;
    
    nodes.forEach(node => {
        if (parent) {
            node.parent = parent;
        }
        
        if (node.children && Array.isArray(node.children)) {
            setupParentReferences(node.children, node);
        }
    });
}

// Utility function to determine if we're in automatic list mode
function isAutomaticListMode(sheetData) {
    return sheetData && sheetData.isListTemplate && sheetData.processedData;
}

// Helper function to truncate text with middle ellipsis
function truncateMiddle(text, maxLength = 35) {
  if (!text || text.length <= maxLength) return text;
  
  const start = Math.floor(maxLength / 2);
  const end = Math.ceil(maxLength / 2);
  
  return text.substring(0, start) + '...' + text.substring(text.length - end);
}

// Function to synchronize row heights across columns in list template
function synchronizeListTemplateRowHeights(sheetContainer) {
    console.log('🔄 Synchronizing row heights across columns for zebra stripe consistency...');
    
    const a4Pages = sheetContainer.querySelectorAll('.a4-page');
    if (!a4Pages.length) {
        console.log('No A4 pages found for row height synchronization');
        return;
    }
    
    a4Pages.forEach((page, pageIndex) => {
        console.log(`📄 Processing page ${pageIndex + 1} for row height sync`);
        
        // Find all level-0 nodes (columns) in this page - handle both direct and nested structures
        let columns = page.querySelectorAll(':scope > .level-0-node');
        
        // If no direct columns found, check for columns inside a container (course-based layout)
        if (columns.length === 0) {
            const columnsContainer = page.querySelector('[style*="display: flex"][style*="flex-direction: row"]');
            if (columnsContainer) {
                columns = columnsContainer.querySelectorAll(':scope > .level-0-node');
                console.log(`Found columns container with ${columns.length} columns`);
            }
        }
        
        if (columns.length === 0) {
            console.log(`No columns found in page ${pageIndex + 1} - checking all level-0-node elements`);
            columns = page.querySelectorAll('.level-0-node');
            console.log(`Found ${columns.length} level-0-node elements anywhere in page`);
        }
        
        if (columns.length === 0) {
            console.log(`Still no columns found in page ${pageIndex + 1}`);
            return;
        }
        
        console.log(`Found ${columns.length} columns in page ${pageIndex + 1}`);
        
        // Find the maximum number of rows across all columns
        let maxRows = 0;
        const columnRowArrays = [];
        
        columns.forEach((column, colIndex) => {
            const level0Children = column.querySelector('.level-0-children');
            if (!level0Children) {
                columnRowArrays.push([]);
                console.log(`Column ${colIndex} has no level-0-children`);
                return;
            }
            
            const rows = level0Children.querySelectorAll(':scope > .level-1-node');
            columnRowArrays.push(Array.from(rows));
            maxRows = Math.max(maxRows, rows.length);
            console.log(`Column ${colIndex} (${column.getAttribute('data-column-name') || 'unknown'}) has ${rows.length} rows`);
        });
        
        console.log(`Maximum rows in page ${pageIndex + 1}: ${maxRows}`);
        
        // First, reset all heights to auto
        for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
            const rowNodes = [];
            
            // Collect all nodes at this row index from all columns
            columnRowArrays.forEach((columnRows, colIndex) => {
                if (columnRows[rowIndex]) {
                    rowNodes.push(columnRows[rowIndex]);
                }
            });
            
            if (rowNodes.length === 0) continue;
            
            // Reset heights to auto to get natural heights
            rowNodes.forEach(node => {
                node.style.height = 'auto';
                node.style.minHeight = 'auto';
            });
        }
        
        // Force a layout recalculation
        page.offsetHeight;
        
        // Use requestAnimationFrame to ensure layout is complete before measuring
        requestAnimationFrame(() => {
            // Now measure and apply synchronized heights
            for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
                const rowNodes = [];
                
                // Collect all nodes at this row index from all columns
                columnRowArrays.forEach((columnRows, colIndex) => {
                    if (columnRows[rowIndex]) {
                        rowNodes.push(columnRows[rowIndex]);
                    }
                });
                
                if (rowNodes.length === 0) continue;
                
                // Find the maximum natural height
                let maxHeight = 0;
                rowNodes.forEach(node => {
                    const rect = node.getBoundingClientRect();
                    const height = rect.height;
                    maxHeight = Math.max(maxHeight, height);
                });
                
                // Apply the maximum height to all nodes in this row
                if (maxHeight > 0) {
                    rowNodes.forEach(node => {
                        node.style.height = `${Math.ceil(maxHeight)}px`;
                        node.style.minHeight = `${Math.ceil(maxHeight)}px`;
                    });
                    console.log(`Row ${rowIndex} in page ${pageIndex + 1}: Set height to ${Math.ceil(maxHeight)}px for ${rowNodes.length} nodes`);
                }
            }
            
            console.log(`✅ Applied synchronized heights for page ${pageIndex + 1}`);
        });
        
        console.log(`✅ Completed row height synchronization for page ${pageIndex + 1}`);
    });
    
    console.log('✅ Row height synchronization completed for all pages');
    debugLogActualPageHeights(sheetContainer);
}

// **ECHTE DYNAMISCHE BEREKENING** - Geen artificiele limieten meer!
const PAGE_CONFIG = {
    // GEEN maxRowsPerPage meer - we berekenen dit dynamisch!
    availableHeight: 1008      // A4 beschikbare ruimte: 297mm ≈ 1122px minus 2×15mm padding ≈ 114px = 1008px
};

// OVERFLOW DETECTION met echte row height sync EN identieke DOM-structuur EN geforceerde CSS
function calculateRowsPerPage(columns, isSingleCourse = false) {
    console.log('🖨️ TRUE SYNCED PAGINATION: test met echte row height sync & DOM structuur & geforceerde CSS...');
    console.log('🖨️ Layout type:', isSingleCourse ? 'SINGLE COURSE (direct columns)' : 'MULTI COURSE (columns container)');
    if (!columns || !Array.isArray(columns) || columns.length === 0) {
        console.log('🖨️ EARLY RETURN: No columns, returning fallback: 1');
        return 1; // Minimale fallback
    }
    const totalRows = columns[0]?.children?.length || 0;
    console.log('🖨️ Total rows found:', totalRows);
    if (totalRows === 0) {
        console.log('🖨️ EARLY RETURN: No rows, returning fallback: 1');
        return 1; // Minimale fallback
    }
    // Maak testpagina
    const testPage = document.createElement('div');
    testPage.style.cssText = `
        position: absolute;
        top: -9999px;
        left: -9999px;
        visibility: hidden;
        width: 210mm;
        height: 297mm;
        box-sizing: border-box;
        padding: 15mm;
        display: flex;
        flex-direction: ${isSingleCourse ? 'row' : 'column'};
        overflow: hidden;
        background: white;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif;
        font-size: 14px;
        color: #333;
    `;
    testPage.className = 'a4-page';
    const sheetContainer = document.createElement('div');
    sheetContainer.className = 'sheet-container';
    sheetContainer.setAttribute('data-is-list-template', 'true');
    
    // CRITICAL: Set correct data-has-courses attribute based on layout type
    if (isSingleCourse) {
        // Single course: no courses
        sheetContainer.removeAttribute('data-has-courses');
    } else {
        // Multi course: has courses
        sheetContainer.setAttribute('data-has-courses', 'true');
    }
    
    sheetContainer.style.cssText = `
        width: 100%;
        height: 100%;
        display: flex;
        flex-direction: column;
        background: white;
        box-sizing: border-box;
    `;
    document.body.appendChild(testPage);
    testPage.appendChild(sheetContainer);
    try {
        let columnsContainer;
        
        if (isSingleCourse) {
            // Single course: columns direct in sheetContainer (geen wrapper)
            columnsContainer = sheetContainer;
        } else {
            // Multi-course: columns in wrapper container
            columnsContainer = document.createElement('div');
            columnsContainer.style.cssText = `
                display: flex;
                flex-direction: row;
                gap: 0;
                flex: 1;
                width: 100%;
                background: white;
            `;
            sheetContainer.appendChild(columnsContainer);
        }
        
        columns.forEach((column, colIndex) => {
            const tempCol = document.createElement('div');
            if (isSingleCourse) {
                // Single course styling - MATCH EXACT CSS: .a4-page > .level-0-node
                let flexValue = '1'; // default
                if (colIndex === 0) flexValue = '1'; // first: 40%
                else if (colIndex === 1) flexValue = '1'; // middle: 20%  
                else if (colIndex === 2) flexValue = '3'; // last: 40%
                
                tempCol.style.cssText = `
                    margin: 0 !important;
                    height: auto !important;
                    display: flex !important;
                    flex-direction: column !important;
                    border-radius: 8px !important;
                    background: white !important;
                    flex: ${flexValue} !important;
                    box-sizing: border-box;
                `;
            } else {
                // Multi-course styling (columns in container)
                tempCol.style.cssText = `
                    flex: 1;
                    display: flex;
                    flex-direction: column;
                    background: white;
                    box-sizing: border-box;
                `;
            }
            tempCol.className = 'level-0-node expanded';
            // --- BELANGRIJK: maak .level-0-children container ---
            const childrenContainer = document.createElement('div');
            childrenContainer.className = 'level-0-children';
            childrenContainer.style.cssText = `
                display: flex;
                flex-direction: column;
                width: 100%;
                background: white;
                box-sizing: border-box;
            `;
            for (let i = 0; i < totalRows; i++) {
                const rowNode = document.createElement('div');
                rowNode.style.cssText = `
                    flex-shrink: 0;
                    box-sizing: border-box;
                    display: flex;
                    flex-direction: column;
                    background: white;
                    margin: 0 !important;
                    border: none !important;
                    border-radius: 0 !important;
                `;
                rowNode.className = 'level-1-node';
                const content = document.createElement('div');
                content.style.cssText = `
                    flex: 1;
                    display: flex;
                    align-items: center;
                    min-height: 100%;
                    padding: 10px 15px !important;
                    margin: 0 !important;
                    border: none !important;
                    border-radius: 0 !important;
                    word-wrap: break-word;
                    overflow-wrap: break-word;
                    background: white;
                    color: #333;
                    font-size: 14px !important;
                    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif;
                    box-sizing: border-box;
                `;
                content.className = 'level-1-content';
                const cellValue = column.children && column.children[i] ? (column.children[i].value || '') : '';
                content.textContent = cellValue;
                rowNode.appendChild(content);
                childrenContainer.appendChild(rowNode);
            }
            tempCol.appendChild(childrenContainer);
            columnsContainer.appendChild(tempCol);
        });
        
        if (!isSingleCourse) {
            // Multi-course: columnsContainer was al toegevoegd aan sheetContainer
        }
        // Single course: columns zijn direct toegevoegd aan sheetContainer (=columnsContainer)
        testPage.offsetHeight;
        // Bereken beschikbare hoogte: 297mm ≈ 1122px, minus 2×15mm padding ≈ 2×57px = 114px
        const pageHeight = 900; // Werkende waarde voor stable pagination
        // Functie die row height sync uitvoert op de eerste N rows
        function syncRows(maxRows) {
            const columnElements = isSingleCourse ? sheetContainer.children : columnsContainer.children;
            let totalHeight = 0;
            for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
                const rowNodes = [];
                for (let colIndex = 0; colIndex < columnElements.length; colIndex++) {
                    // --- BELANGRIJK: pak rows uit .level-0-children ---
                    const childrenContainer = columnElements[colIndex].querySelector('.level-0-children');
                    if (!childrenContainer) continue;
                    const rowNode = childrenContainer.children[rowIndex];
                    if (rowNode) rowNodes.push(rowNode);
                }
                if (rowNodes.length === 0) continue;
                // Meet max hoogte
                let maxHeight = 0;
                rowNodes.forEach(node => {
                    node.style.height = '';
                    node.style.minHeight = '';
                });
                testPage.offsetHeight;
                rowNodes.forEach(node => {
                    const rect = node.getBoundingClientRect();
                    const height = rect.height;
                    maxHeight = Math.max(maxHeight, height);
                });
                const syncHeight = Math.ceil(maxHeight);
                rowNodes.forEach(node => {
                    node.style.height = `${syncHeight}px`;
                    node.style.minHeight = `${syncHeight}px`;
                });
                totalHeight += syncHeight;
            }
            return totalHeight;
        }
        // **VEILIGHEIDS-STRATEGIE**: Test met zwaarste content uit de dataset
        // Zoek de langste content om een conservatieve schatting te maken
        let maxContentLength = 0;
        let heaviestRowIndex = 0;
        for (let i = 0; i < totalRows; i++) {
            for (let colIndex = 0; colIndex < columns.length; colIndex++) {
                const cellValue = columns[colIndex].children && columns[colIndex].children[i] 
                    ? (columns[colIndex].children[i].value || '') : '';
                if (cellValue.length > maxContentLength) {
                    maxContentLength = cellValue.length;
                    heaviestRowIndex = i;
                }
            }
        }
        console.log(`🖨️ SAFETY: Found heaviest content at row ${heaviestRowIndex} with ${maxContentLength} chars`);
        
        // Pas de test-content aan om zwaarste content te gebruiken
        const testColumnElements = isSingleCourse ? sheetContainer.children : columnsContainer.children;
        for (let colIndex = 0; colIndex < testColumnElements.length; colIndex++) {
            const childrenContainer = testColumnElements[colIndex].querySelector('.level-0-children');
            if (!childrenContainer) continue;
            
            // Zet zwaarste content in alle test-rows voor conservatieve schatting
            if (heaviestRowIndex < columns.length) {
                const heaviestContent = columns[Math.min(colIndex, columns.length-1)].children && 
                    columns[Math.min(colIndex, columns.length-1)].children[heaviestRowIndex] 
                    ? (columns[Math.min(colIndex, columns.length-1)].children[heaviestRowIndex].value || '') : 'Test content';
                
                for (let rowIndex = 0; rowIndex < childrenContainer.children.length; rowIndex++) {
                    const rowContent = childrenContainer.children[rowIndex].querySelector('.level-1-content');
                    if (rowContent) {
                        rowContent.textContent = heaviestContent;
                    }
                }
            }
        }
        
        // Binaire search
        let low = 1;
        let high = totalRows; // Geen artificiele limiet meer
        let bestFit = low;
        while (low <= high) {
            // Reset alle rows zichtbaar en heights leeg
            const columnElements = isSingleCourse ? sheetContainer.children : columnsContainer.children;
            for (let colIndex = 0; colIndex < columnElements.length; colIndex++) {
                const childrenContainer = columnElements[colIndex].querySelector('.level-0-children');
                if (!childrenContainer) continue;
                for (let rowIndex = 0; rowIndex < childrenContainer.children.length; rowIndex++) {
                    const row = childrenContainer.children[rowIndex];
                    row.style.display = '';
                    row.style.height = '';
                    row.style.minHeight = '';
                }
            }
            testPage.offsetHeight;
            const mid = Math.floor((low + high) / 2);
            // Hide rows > mid
            const columnElementsForHiding = isSingleCourse ? sheetContainer.children : columnsContainer.children;
            for (let colIndex = 0; colIndex < columnElementsForHiding.length; colIndex++) {
                const childrenContainer = columnElementsForHiding[colIndex].querySelector('.level-0-children');
                if (!childrenContainer) continue;
                for (let rowIndex = 0; rowIndex < childrenContainer.children.length; rowIndex++) {
                    const row = childrenContainer.children[rowIndex];
                    if (rowIndex >= mid) row.style.display = 'none';
                }
            }
            testPage.offsetHeight;
            const testHeight = syncRows(mid);
            console.log(`   🔍 Test ${mid} rows, synced height: ${testHeight}px (max: ${pageHeight}px)`);
            // **KLEINE VEILIGHEIDSMARGIN**: 5% extra ruimte om overflows te voorkomen
            const safetyMargin = pageHeight * 0.95; // 95% van beschikbare ruimte = 5% margin
            if (testHeight <= safetyMargin) {
                bestFit = mid;
                low = mid + 1;
            } else {
                high = mid - 1;
            }
        }
        console.log(`   🎯 TRUE SYNCED FIT: ${bestFit} rows per page`);
        console.log(`🖨️ FUNCTION RETURNING: ${bestFit} rows per page for ${isSingleCourse ? 'SINGLE' : 'MULTI'} course`);
        return bestFit;
    } finally {
        if (testPage.parentNode) {
            document.body.removeChild(testPage);
        }
    }
}

// NIEUWE FUNCTIE: Meet echte rijhoogtes en bepaal page breaks met titel ruimte
function calculateRealPageBreaksWithTitle(columns, maxPageHeight, titleHeight, isSingleCourse = true) {
    console.log(`🎯 Calculating page breaks with title height: ${titleHeight}px`);
    
    const totalRows = columns[0] && columns[0].children ? columns[0].children.length : 0;
    if (totalRows === 0) {
        console.log('❌ No rows to calculate page breaks for');
        return [0];
    }
    
    // Create temporary measurement page
    const testPage = document.createElement('div');
    testPage.style.cssText = `
        position: fixed;
        top: -9999px;
        left: -9999px;
        width: 210mm;
        height: 297mm;
        padding: 15mm;
        box-sizing: border-box;
        display: flex;
        flex-direction: column;
        visibility: hidden;
        pointer-events: none;
    `;
    
    const sheetContainer = document.createElement('div');
    sheetContainer.style.cssText = `
        display: flex;
        flex-direction: ${isSingleCourse ? 'row' : 'column'};
        width: 100%;
        height: 100%;
        gap: 0;
    `;
    testPage.appendChild(sheetContainer);
    
    document.body.appendChild(testPage);
    
    try {
        let columnsContainer;
        
        if (isSingleCourse) {
            columnsContainer = sheetContainer;
        } else {
            // Multi-course: Simulate course header + columns layout like real pages
            const courseHeader = document.createElement('div');
            courseHeader.style.cssText = `
                background: #34a3d7;
                color: white;
                padding: 8px 16px;
                border-radius: 50px;
                font-size: 1rem;
                font-weight: bold;
                width: fit-content;
                margin-bottom: 0.5rem;
                flex-shrink: 0;
            `;
            courseHeader.textContent = 'Test Course';
            sheetContainer.appendChild(courseHeader);
            
            columnsContainer = document.createElement('div');
            columnsContainer.style.cssText = `
                display: flex;
                flex-direction: row;
                gap: 0;
                flex: 1;
                width: 100%;
                background: white;
            `;
            sheetContainer.appendChild(columnsContainer);
        }
        
        // Render alle kolommen met echte content
        columns.forEach((column, columnIndex) => {
            const tempCol = document.createElement('div');
            tempCol.className = 'level-0-node';
            tempCol.style.cssText = `
                flex: ${columnIndex === 2 ? '3' : '1'};
                display: flex;
                flex-direction: column;
                min-width: 0;
                background: white;
                border-radius: 8px;
                padding: 0;
                margin: 0;
            `;
            
            const header = document.createElement('div');
            header.className = 'level-0-header';
            header.style.cssText = `
                background: #34a3d7;
                color: white;
                padding: 8px 12px;
                font-weight: bold;
                text-align: center;
                border-radius: 8px 8px 0 0;
            `;
            header.textContent = column.value || `Column ${columnIndex + 1}`;
            tempCol.appendChild(header);
            
            const childrenContainer = document.createElement('div');
            childrenContainer.className = 'level-0-children';
            childrenContainer.style.cssText = `
                flex: 1;
                display: flex;
                flex-direction: column;
                background: white;
                border-radius: 0 0 8px 8px;
                overflow: hidden;
            `;
            
            if (column.children) {
                column.children.forEach((child, childIndex) => {
                    const childElement = document.createElement('div');
                    childElement.className = 'level-1-node';
                    childElement.style.cssText = `
                        padding: 8px 12px;
                        border-bottom: 1px solid #eee;
                        background: ${childIndex % 2 === 0 ? 'white' : '#f8f9fa'};
                        font-size: 14px;
                        line-height: 1.4;
                        min-height: 20px;
                        display: flex;
                        align-items: center;
                    `;
                    childElement.textContent = child.value || '';
                    childrenContainer.appendChild(childElement);
                });
            }
            
            tempCol.appendChild(childrenContainer);
            columnsContainer.appendChild(tempCol);
        });
        
        testPage.offsetHeight;
        synchronizeTestRowHeights(isSingleCourse ? sheetContainer : columnsContainer);
        
        // Nu meet de echte hoogtes en bepaal page breaks met titel ruimte
        const pageBreaks = [];
        let currentPageHeight = 0;
        let currentRowStart = 0;
        let isFirstPage = true;
        
        const firstColumn = (isSingleCourse ? sheetContainer : columnsContainer).children[0];
        const firstChildrenContainer = firstColumn.querySelector('.level-0-children');
        
        for (let rowIndex = 0; rowIndex < totalRows; rowIndex++) {
            const rowElement = firstChildrenContainer.children[rowIndex];
            const rowHeight = rowElement.getBoundingClientRect().height;
            
            console.log(`📏 Row ${rowIndex}: ${Math.ceil(rowHeight)}px`);
            
            // Bepaal beschikbare hoogte (minder voor eerste pagina vanwege titel)
            const availableHeight = isFirstPage ? (maxPageHeight - titleHeight) : maxPageHeight;
            console.log(`📐 Available height for ${isFirstPage ? 'first' : 'subsequent'} page: ${availableHeight}px`);
            
            if (currentPageHeight + rowHeight > availableHeight && currentPageHeight > 0) {
                pageBreaks.push(rowIndex);
                console.log(`📄 Page break at row ${rowIndex}, previous page height: ${Math.ceil(currentPageHeight)}px`);
                currentPageHeight = rowHeight;
                isFirstPage = false; // Volgende pagina's hebben geen titel meer
            } else {
                currentPageHeight += rowHeight;
            }
        }
        
        pageBreaks.push(totalRows);
        console.log(`✅ Final page breaks: [${pageBreaks.join(', ')}]`);
        
        return pageBreaks;
        
    } finally {
        document.body.removeChild(testPage);
    }
}

// NIEUWE FUNCTIE: Meet echte rijhoogtes en bepaal page breaks
function calculateRealPageBreaks(columns, maxPageHeight, isSingleCourse = true) {
    console.log('📏 REAL MEASUREMENT: Creating test render to measure actual row heights...');
    console.log('📏 Layout type:', isSingleCourse ? 'SINGLE COURSE' : 'MULTI COURSE');
    
    if (!columns || !Array.isArray(columns) || columns.length === 0) {
        return [1]; // fallback
    }
    
    const totalRows = columns[0]?.children?.length || 0;
    if (totalRows === 0) {
        return [1]; // fallback
    }
    
    // Render alle rijen in test-container met echte styling
    const testPage = document.createElement('div');
    testPage.className = 'a4-page';
    testPage.style.cssText = `
        position: absolute;
        top: -9999px;
        left: -9999px;
        visibility: hidden;
        width: 210mm;
        height: auto; /* Laat hoogte vrij groeien */
        box-sizing: border-box;
        padding: 15mm;
        display: flex;
        flex-direction: column; /* BOTH use column to simulate real A4 page layout */
        background: white;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif;
        font-size: 14px;
        color: #333;
    `;
    
    const sheetContainer = document.createElement('div');
    sheetContainer.className = 'sheet-container';
    sheetContainer.setAttribute('data-is-list-template', 'true');
    
    // Set correct data-has-courses attribute
    if (isSingleCourse) {
        // Single course = GEEN data-has-courses attribuut
        sheetContainer.removeAttribute('data-has-courses');
    } else {
        // Multi course = WEL data-has-courses attribuut
        sheetContainer.setAttribute('data-has-courses', 'true');
    }
    
    testPage.appendChild(sheetContainer);
    document.body.appendChild(testPage);
    
    try {
        let columnsContainer;
        
        if (isSingleCourse) {
            // Single course: columns direct in sheetContainer
            columnsContainer = sheetContainer;
        } else {
            // Multi-course: Simulate course header + columns layout like real pages
            const courseHeader = document.createElement('div');
            courseHeader.style.cssText = `
                background: #34a3d7;
                color: white;
                padding: 8px 16px;
                border-radius: 50px;
                font-size: 1rem;
                font-weight: bold;
                width: fit-content;
                margin-bottom: 0.5rem;
                flex-shrink: 0;
            `;
            courseHeader.textContent = 'Test Course';
            sheetContainer.appendChild(courseHeader);
            
            // Multi-course: columns in wrapper container
            columnsContainer = document.createElement('div');
            columnsContainer.style.cssText = `
                display: flex;
                flex-direction: row;
                gap: 0;
                flex: 1;
                width: 100%;
                background: white;
            `;
            sheetContainer.appendChild(columnsContainer);
        }
        
        // Render alle kolommen met echte content
        columns.forEach((column, colIndex) => {
            const tempCol = document.createElement('div');
            tempCol.className = 'level-0-node expanded';
            
            if (isSingleCourse) {
                // Single course kolom styling met correcte flex ratios
                let flexValue = '1';
                if (colIndex === 0) flexValue = '1'; // first: 40%
                else if (colIndex === 1) flexValue = '1'; // middle: 20%  
                else if (colIndex === 2) flexValue = '3'; // last: 40%
                
                tempCol.style.cssText = `
                    margin: 0 !important;
                    height: auto !important;
                    display: flex !important;
                    flex-direction: column !important;
                    border-radius: 8px !important;
                    background: white !important;
                    flex: ${flexValue} !important;
                    box-sizing: border-box;
                `;
            } else {
                // Multi-course kolom styling
                tempCol.style.cssText = `
                    flex: 1;
                    display: flex;
                    flex-direction: column;
                    background: white;
                    box-sizing: border-box;
                `;
            }
            
            const childrenContainer = document.createElement('div');
            childrenContainer.className = 'level-0-children';
            childrenContainer.style.cssText = `
                display: flex;
                flex-direction: column;
                width: 100%;
                background: white;
                box-sizing: border-box;
            `;
            
            // Render alle rijen met echte content
            for (let i = 0; i < totalRows; i++) {
                const rowNode = document.createElement('div');
                rowNode.className = 'level-1-node';
                rowNode.style.cssText = `
                    flex-shrink: 0;
                    box-sizing: border-box;
                    display: flex;
                    flex-direction: column;
                    background: white;
                    margin: 0 !important;
                    border: none !important;
                    border-radius: 0 !important;
                `;
                
                const content = document.createElement('div');
                content.className = 'level-1-content';
                content.style.cssText = `
                    flex: 1;
                    display: flex;
                    align-items: center;
                    min-height: 100%;
                    padding: 10px 15px !important;
                    margin: 0 !important;
                    border: none !important;
                    border-radius: 0 !important;
                    word-wrap: break-word;
                    overflow-wrap: break-word;
                    background: white;
                    color: #333;
                    font-size: 14px !important;
                    font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, 'Open Sans', 'Helvetica Neue', sans-serif;
                    box-sizing: border-box;
                `;
                
                // Echte content uit de data
                const cellValue = column.children && column.children[i] ? (column.children[i].value || '') : '';
                content.textContent = cellValue;
                
                rowNode.appendChild(content);
                childrenContainer.appendChild(rowNode);
            }
            
            tempCol.appendChild(childrenContainer);
            columnsContainer.appendChild(tempCol);
        });
        
        // Force layout
        testPage.offsetHeight;
        
        // Synchroniseer rijhoogtes (zoals in echte rendering)
        synchronizeTestRowHeights(isSingleCourse ? sheetContainer : columnsContainer);
        
        // Nu meet de echte hoogtes en bepaal page breaks
        const pageBreaks = [];
        let currentPageHeight = 0;
        let currentRowStart = 0;
        
        const firstColumn = (isSingleCourse ? sheetContainer : columnsContainer).children[0];
        const firstChildrenContainer = firstColumn.querySelector('.level-0-children');
        
        for (let rowIndex = 0; rowIndex < totalRows; rowIndex++) {
            const rowElement = firstChildrenContainer.children[rowIndex];
            const rowHeight = rowElement.getBoundingClientRect().height;
            
            console.log(`📏 Row ${rowIndex}: ${Math.ceil(rowHeight)}px`);
            
            // Check of deze rij nog past op huidige pagina
            if (currentPageHeight + rowHeight > maxPageHeight && currentPageHeight > 0) {
                // Start nieuwe pagina
                pageBreaks.push(rowIndex);
                console.log(`📄 Page break at row ${rowIndex}, previous page height: ${Math.ceil(currentPageHeight)}px`);
                currentPageHeight = rowHeight; // Start nieuwe pagina met deze rij
            } else {
                currentPageHeight += rowHeight;
            }
        }
        
        // Laatste pagina
        pageBreaks.push(totalRows);
        console.log(`📄 Final page ends at row ${totalRows}, height: ${Math.ceil(currentPageHeight)}px`);
        
        return pageBreaks;
        
    } finally {
        if (testPage.parentNode) {
            document.body.removeChild(testPage);
        }
    }
}

// Helper functie voor row height sync in test
function synchronizeTestRowHeights(container) {
    const columns = Array.from(container.children);
    const maxRows = Math.max(...columns.map(col => {
        const childrenContainer = col.querySelector('.level-0-children');
        return childrenContainer ? childrenContainer.children.length : 0;
    }));
    
    for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
        const rowNodes = [];
        
        columns.forEach(column => {
            const childrenContainer = column.querySelector('.level-0-children');
            if (childrenContainer && childrenContainer.children[rowIndex]) {
                rowNodes.push(childrenContainer.children[rowIndex]);
            }
        });
        
        if (rowNodes.length === 0) continue;
        
        // Reset heights
        rowNodes.forEach(node => {
            node.style.height = 'auto';
            node.style.minHeight = 'auto';
        });
        
        // Force layout
        container.offsetHeight;
        
        // Find max height
        let maxHeight = 0;
        rowNodes.forEach(node => {
            const rect = node.getBoundingClientRect();
            maxHeight = Math.max(maxHeight, rect.height);
        });
        
        // Apply synchronized height
        const syncHeight = Math.ceil(maxHeight);
        rowNodes.forEach(node => {
            node.style.height = `${syncHeight}px`;
            node.style.minHeight = `${syncHeight}px`;
        });
    }
}

// Function to hide repeating blok names per page (show first occurrence per page)
function hideRepeatingBlokNamesPerPage(sheetContainer) {
    console.log('🔄 Hiding repeating blok names per page...');
    
    const a4Pages = sheetContainer.querySelectorAll('.a4-page');
    if (!a4Pages.length) {
        console.log('No A4 pages found for blok name hiding');
        return;
    }
    
    a4Pages.forEach((page, pageIndex) => {
        console.log(`📄 Processing page ${pageIndex + 1} for blok name hiding`);
        
        // Find blok column (first column)
        let blokColumn = page.querySelector(':scope > .level-0-node');
        
        // If no direct columns found, check for columns inside a container (course-based layout)
        if (!blokColumn) {
            const columnsContainer = page.querySelector('[style*="display: flex"][style*="flex-direction: row"]');
            if (columnsContainer) {
                blokColumn = columnsContainer.querySelector(':scope > .level-0-node');
            }
        }
        
        if (!blokColumn) {
            console.log(`No blok column found in page ${pageIndex + 1}`);
            return;
        }
        
        // Find all blok content elements in this page
        const level0Children = blokColumn.querySelector('.level-0-children');
        if (!level0Children) {
            console.log(`No level-0-children found in blok column for page ${pageIndex + 1}`);
            return;
        }
        
        const blokContentElements = level0Children.querySelectorAll(':scope > .level-1-node .level-1-content');
        if (!blokContentElements.length) {
            console.log(`No blok content elements found for page ${pageIndex + 1}`);
            return;
        }
        
        console.log(`Found ${blokContentElements.length} blok content elements in page ${pageIndex + 1}`);
        
        // Track seen blok values within THIS PAGE ONLY
        const seenBlokValuesThisPage = new Set();
        
        blokContentElements.forEach((contentElement, index) => {
            const currentBlokValue = contentElement.textContent.trim();
            
            if (currentBlokValue === '' || currentBlokValue === 'undefined') {
                // Skip empty values
                return;
            }
            
            if (seenBlokValuesThisPage.has(currentBlokValue)) {
                // This is a repeat within this page - hide text but keep background
                contentElement.style.color = 'transparent';
                console.log(`Hidden repeat blok "${currentBlokValue}" at index ${index} on page ${pageIndex + 1}`);
            } else {
                // This is the first occurrence on this page - make sure it's visible
                contentElement.style.color = '';
                seenBlokValuesThisPage.add(currentBlokValue);
                console.log(`Showing first occurrence of blok "${currentBlokValue}" at index ${index} on page ${pageIndex + 1}`);
            }
        });
        
        console.log(`✅ Completed blok name hiding for page ${pageIndex + 1}. Unique bloks on this page: ${seenBlokValuesThisPage.size}`);
    });
    
    console.log('✅ Blok name hiding completed for all pages');
}

// Setup resize listener for row height synchronization
function setupRowHeightSyncOnResize() {
    let resizeTimeout;
    
    window.addEventListener('resize', () => {
        clearTimeout(resizeTimeout);
        resizeTimeout = setTimeout(() => {
            const activeSheetContainer = document.querySelector('.sheet-content.active .sheet-container[data-is-list-template="true"]');
            if (activeSheetContainer) {
                console.log('🔄 Window resized - re-synchronizing row heights...');
                synchronizeListTemplateRowHeights(activeSheetContainer);
                hideRepeatingBlokNamesPerPage(activeSheetContainer);
            }
        }, 250); // Debounce resize events
    });
}

// Initialize resize handler when module loads
if (typeof window !== 'undefined') {
    setupRowHeightSyncOnResize();
    
    // Add global debug function for manual testing
    window.debugRowHeightSync = function() {
        const activeSheetContainer = document.querySelector('.sheet-content.active .sheet-container[data-is-list-template="true"]');
        if (activeSheetContainer) {
            console.log('🔧 Manual row height synchronization triggered...');
            synchronizeListTemplateRowHeights(activeSheetContainer);
            hideRepeatingBlokNamesPerPage(activeSheetContainer);
        } else {
            console.log('❌ No active list template sheet found');
        }
    };
}

function debugLogActualPageHeights(sheetContainer) {
    const a4Pages = sheetContainer.querySelectorAll('.a4-page');
    a4Pages.forEach((page, pageIndex) => {
        let totalHeight = 0;
        let rowHeights = [];
        // Vind alle kolommen
        let columns = page.querySelectorAll(':scope > .level-0-node');
        if (columns.length === 0) {
            const columnsContainer = page.querySelector('[style*="display: flex"][style*="flex-direction: row"]');
            if (columnsContainer) {
                columns = columnsContainer.querySelectorAll(':scope > .level-0-node');
            }
        }
        if (columns.length === 0) {
            columns = page.querySelectorAll('.level-0-node');
        }
        if (columns.length === 0) return;
        // Pak rows uit .level-0-children
        const maxRows = Array.from(columns).reduce((max, col) => {
            const children = col.querySelector('.level-0-children');
            return Math.max(max, children ? children.children.length : 0);
        }, 0);
        for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
            const rowNodes = [];
            columns.forEach(col => {
                const children = col.querySelector('.level-0-children');
                if (children && children.children[rowIndex]) {
                    rowNodes.push(children.children[rowIndex]);
                }
            });
            if (rowNodes.length === 0) continue;
            let maxHeight = 0;
            rowNodes.forEach(node => {
                const rect = node.getBoundingClientRect();
                const height = rect.height;
                maxHeight = Math.max(maxHeight, height);
            });
            rowHeights.push(maxHeight);
            totalHeight += maxHeight;
        }
        console.log(`\uD83D\uDCC4 [DEBUG] Page ${pageIndex + 1}: total synced height = ${totalHeight.toFixed(1)}px, rows = ${rowHeights.length}`);
        rowHeights.forEach((h, i) => {
            console.log(`    Row ${i}: ${h.toFixed(1)}px`);
        });
    });
}

// Export functions for other modules
window.ExcelViewerSheetManager = {
  loadSheetData,
  saveHierarchyConfiguration,
  saveCurrentHierarchyConfiguration,
  renderSheet,
  truncateMiddle,
  setupParentReferences: setupParentReferences,
  isAutomaticListMode: isAutomaticListMode,
  synchronizeListTemplateRowHeights: synchronizeListTemplateRowHeights,
  hideRepeatingBlokNamesPerPage: hideRepeatingBlokNamesPerPage
}; 