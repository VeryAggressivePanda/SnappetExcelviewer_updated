document.addEventListener('DOMContentLoaded', function() {
  // DOM Elements
  const tabButtons = document.querySelectorAll('.tab-button');
  const sheetContents = document.querySelectorAll('.sheet-content');
  const fontSizeSelect = document.getElementById('font-size');
  const themeSelect = document.getElementById('theme');
  const cellPaddingSelect = document.getElementById('cell-padding');
  const showGridCheckbox = document.getElementById('show-grid');
  
  // Tab switching with lazy loading
  tabButtons.forEach(button => {
    button.addEventListener('click', async () => {
      const sheetId = button.getAttribute('data-sheet-id');
      
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
          }
        } else {
          content.classList.remove('active');
        }
      });
    });
  });
  
  // Load sheet data via AJAX
  async function loadSheetData(sheetId) {
    try {
      const sheetContent = document.getElementById(`sheet-${sheetId}`);
      
      // Show loading indicator
      sheetContent.innerHTML = '<div class="loading-indicator"><p>Loading sheet data...</p></div>';
      
      // Fetch the sheet data
      const response = await fetch(`/api/sheet/${window.excelData.fileId}/${sheetId}`);
      
      if (!response.ok) {
        throw new Error(`Failed to load sheet data: ${response.statusText}`);
      }
      
      const result = await response.json();
      console.log('API response:', result);
      
      // Check the structure of the response
      if (!result.data) {
        console.error('Missing data property in API response:', result);
        sheetContent.innerHTML = '<div class="empty-sheet"><p>Error: Invalid data received from server</p></div>';
        window.excelData.sheetsLoaded[sheetId] = true;
        return;
      }
      
      // Create the HTML for the hierarchical structure
      const html = renderHierarchicalData(result.data);
      
      // Update the content
      sheetContent.innerHTML = html;
      
      // Set up expand/collapse functionality
      setupNodeInteractions(sheetContent);
      
      // Mark as loaded and store the data
      window.excelData.sheetsLoaded[sheetId] = result.data;
      
      // Populate the export column select
      populateColumnSelects(result.data);
      
      // Update export controls with preview button
      updateExportControls(result.data);
      
      // Apply current style settings
      applyStyles();
      
    } catch (error) {
      console.error('Error loading sheet data:', error);
      const sheetContent = document.getElementById(`sheet-${sheetId}`);
      sheetContent.innerHTML = `<div class="error-message"><p>Error loading sheet data: ${error.message}</p></div>`;
    }
  }
  
  // Function to render hierarchical data
  function renderHierarchicalData(data) {
    console.log('Rendering hierarchical data:', data);
    
    if (!data || !data.headers || !data.root) {
      console.error('Invalid data structure for rendering:', data);
      return '<div class="empty-sheet"><p>No data to display</p></div>';
    }
    
    // Create the container
    const container = document.createElement('div');
    container.className = 'excel-hierarchical-container';
    
    // Add controls at the top
    const controls = document.createElement('div');
    controls.className = 'controls';
    
    const expandAllButton = document.createElement('button');
    expandAllButton.className = 'control-button expand-all-button';
    expandAllButton.textContent = 'Expand All';
    
    const collapseAllButton = document.createElement('button');
    collapseAllButton.className = 'control-button collapse-all-button';
    collapseAllButton.textContent = 'Collapse All';
    
    const searchGroup = document.createElement('div');
    searchGroup.className = 'search-group';
    
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.className = 'search-input';
    searchInput.placeholder = 'Search...';
    
    searchGroup.appendChild(searchInput);
    controls.appendChild(expandAllButton);
    controls.appendChild(collapseAllButton);
    controls.appendChild(searchGroup);
    container.appendChild(controls);
    
    // Check if we have children to render
    if (!data.root.children || data.root.children.length === 0) {
      console.log('No children to render in root node');
      const emptyMessage = document.createElement('div');
      emptyMessage.className = 'excel-hierarchical-content';
      emptyMessage.innerHTML = '<div class="empty-sheet"><p>No hierarchical data to display</p></div>';
      container.appendChild(emptyMessage);
      return container.outerHTML;
    }
    
    // Render the hierarchical content
    console.log(`Found ${data.root.children.length} top-level items to render`);
    const content = document.createElement('div');
    content.className = 'excel-hierarchical-content';
    
    // Render children
    const childrenContainer = renderChildren(data.root.children, 0);
    content.appendChild(childrenContainer);
    container.appendChild(content);
    
    return container.outerHTML;
  }
  
  // Function to render children recursively
  function renderChildren(children, level) {
    const container = document.createElement('div');
    container.className = 'excel-node-children';
    
    children.forEach(child => {
      const node = document.createElement('div');
      node.className = `excel-node level-${level}`;
      
      const header = document.createElement('div');
      header.className = 'excel-node-header';
      
      // Only add expand button for Blok (level 0) and Week (level 1)
      if (level < 2) {
        const expandButton = document.createElement('div');
        expandButton.className = 'node-expand-button';
        header.appendChild(expandButton);
      }
      
      const title = document.createElement('div');
      title.className = 'node-title';
      
      // Get the title based on the level
      let titleText = 'Untitled';
      if (level === 0) {
        titleText = child.value || 'Blok';
      } else if (level === 1) {
        titleText = child.value || 'Week';
      } else if (level === 2) {
        // For Les level, try to get the value from properties first, then from the node value
        const lesProperty = child.properties?.find(prop => prop.columnName === 'Les');
        titleText = lesProperty?.value || child.value || 'Les';
      }
      title.textContent = titleText;
      
      header.appendChild(title);
      node.appendChild(header);
      
      // Add properties
      if (child.properties && child.properties.length > 0) {
        const propertiesContainer = document.createElement('div');
        propertiesContainer.className = 'excel-node-properties';
        
        const propertyRow = document.createElement('div');
        propertyRow.className = 'property-row';
        
        child.properties.forEach(prop => {
          const property = document.createElement('div');
          property.className = 'excel-node-property';
          
          const propertyName = document.createElement('div');
          propertyName.className = 'property-name';
          propertyName.textContent = prop.columnName;
          
          const propertyValue = document.createElement('div');
          propertyValue.className = 'property-value';
          propertyValue.textContent = prop.value || '';
          
          property.appendChild(propertyName);
          property.appendChild(propertyValue);
          propertyRow.appendChild(property);
        });
        
        propertiesContainer.appendChild(propertyRow);
        node.appendChild(propertiesContainer);
      }
      
      // Add children if they exist
      if (child.children && child.children.length > 0) {
        const childrenContainer = renderChildren(child.children, level + 1);
        node.appendChild(childrenContainer);
      }
      
      container.appendChild(node);
    });
    
    return container;
  }
  
  // Font size control
  fontSizeSelect.addEventListener('change', () => {
    document.documentElement.style.setProperty('--font-size', fontSizeSelect.value);
  });
  
  // Theme control
  themeSelect.addEventListener('change', () => {
    // Remove all theme classes
    document.body.classList.remove('dark-theme', 'light-theme', 'colorful-theme');
    
    // Add selected theme class if not default
    if (themeSelect.value !== 'default') {
      document.body.classList.add(`${themeSelect.value}-theme`);
    }
  });
  
  // Cell padding control
  cellPaddingSelect.addEventListener('change', () => {
    document.documentElement.style.setProperty('--cell-padding', cellPaddingSelect.value);
  });
  
  // Grid visibility control
  showGridCheckbox.addEventListener('change', () => {
    applyStyles();
  });
  
  // Apply styles to all cells
  function applyStyles() {
    const excelCells = document.querySelectorAll('.excel-cell');
    
    excelCells.forEach(cell => {
      if (showGridCheckbox.checked) {
        cell.style.border = '1px solid var(--grid-color)';
      } else {
        cell.style.border = 'none';
      }
    });
  }
  
  // Handle window resize for responsive design
  window.addEventListener('resize', () => {
    adjustContentHeight();
  });
  
  // Adjust content height for scrolling
  function adjustContentHeight() {
    const header = document.querySelector('header');
    const footer = document.querySelector('footer');
    const styleControls = document.querySelector('.style-controls');
    const sheetTabs = document.querySelector('.sheet-tabs');
    const sheetContentContainer = document.querySelector('.sheet-content-container');
    
    if (header && footer && styleControls && sheetTabs && sheetContentContainer) {
      const windowHeight = window.innerHeight;
      const headerHeight = header.offsetHeight;
      const footerHeight = footer.offsetHeight;
      const styleControlsHeight = styleControls.offsetHeight;
      const sheetTabsHeight = sheetTabs.offsetHeight;
      const padding = 60; // Additional padding
      
      const availableHeight = windowHeight - headerHeight - footerHeight - styleControlsHeight - sheetTabsHeight - padding;
      
      sheetContentContainer.style.maxHeight = `${availableHeight}px`;
    }
  }
  
  // Initialize
  function initialize() {
    adjustContentHeight();
    applyStyles();
    
    // Initialize excelData if not exists
    if (!window.excelData) {
      window.excelData = { sheetsLoaded: {} };
    }
    
    // Load the first sheet automatically
    const activeTab = document.querySelector('.tab-button.active');
    if (activeTab) {
      const sheetId = activeTab.getAttribute('data-sheet-id');
      if (sheetId) {
        loadSheetData(sheetId);
      }
    }
  }
  
  // Function to setup node interactions
  function setupNodeInteractions(container) {
    // Setup expand/collapse buttons
    const expandAllButton = container.querySelector('.expand-all-button');
    const collapseAllButton = container.querySelector('.collapse-all-button');
    const searchInput = container.querySelector('.search-input');
    
    // Function to toggle node state
    function toggleNodeState(node, forceCollapsed = null) {
      const isLevel0Or1 = node.classList.contains('level-0') || node.classList.contains('level-1');
      if (!isLevel0Or1) return;
      
      const children = node.querySelector('.excel-node-children');
      if (!children) return;
      
      if (forceCollapsed === null) {
        const isCollapsed = node.classList.contains('collapsed');
        node.classList.toggle('collapsed');
        children.style.display = isCollapsed ? 'block' : 'none';
        
        // Force a reflow to ensure the transition works
        node.offsetHeight;
      } else {
        node.classList.toggle('collapsed', forceCollapsed);
        children.style.display = forceCollapsed ? 'none' : 'block';
        
        // Force a reflow to ensure the transition works
        node.offsetHeight;
      }
    }
    
    // Expand all button
    if (expandAllButton) {
      expandAllButton.addEventListener('click', () => {
        const allNodes = container.querySelectorAll('.excel-node');
        allNodes.forEach(node => toggleNodeState(node, false));
      });
    }
    
    // Collapse all button
    if (collapseAllButton) {
      collapseAllButton.addEventListener('click', () => {
        const allNodes = container.querySelectorAll('.excel-node');
        allNodes.forEach(node => toggleNodeState(node, true));
      });
    }
    
    // Setup individual node click handlers
    const allNodes = container.querySelectorAll('.excel-node');
    allNodes.forEach(node => {
      const header = node.querySelector('.excel-node-header');
      const expandButton = node.querySelector('.node-expand-button');
      
      if (header && expandButton) {
        // Header click handler
        header.addEventListener('click', (e) => {
          if (e.target === expandButton) return;
          toggleNodeState(node);
        });
        
        // Expand button click handler
        expandButton.addEventListener('click', (e) => {
          e.stopPropagation();
          toggleNodeState(node);
        });
      }
    });
    
    // Initialize nodes with correct initial state
    const initialNodes = container.querySelectorAll('.excel-node');
    initialNodes.forEach(node => {
      if (node.classList.contains('level-0')) {
        // Level 0 (Blok) nodes start expanded
        toggleNodeState(node, false);
      } else if (node.classList.contains('level-1')) {
        // Level 1 (Week) nodes start collapsed
        toggleNodeState(node, true);
      }
    });
    
    // Setup search functionality
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        const searchTerm = e.target.value.toLowerCase();
        const nodes = container.querySelectorAll('.excel-node');
        
        if (searchTerm.length > 2) {
          nodes.forEach(node => {
            const nodeTitle = node.querySelector('.node-title');
            const nodeProperties = node.querySelectorAll('.property-value');
            
            let found = nodeTitle && nodeTitle.textContent.toLowerCase().includes(searchTerm);
            
            if (!found) {
              nodeProperties.forEach(prop => {
                if (prop.textContent.toLowerCase().includes(searchTerm)) {
                  found = true;
                }
              });
            }
            
            if (found) {
              node.style.display = '';
              let parent = node.parentElement.closest('.excel-node');
              while (parent) {
                toggleNodeState(parent, false);
                parent = parent.parentElement.closest('.excel-node');
              }
            } else {
              node.style.display = 'none';
            }
          });
        } else {
          nodes.forEach(node => {
            node.style.display = '';
            if (node.classList.contains('collapsed')) {
              toggleNodeState(node, true);
            }
          });
        }
      });
    }
  }
  
  // Add event listeners for PDF export
  document.getElementById('export-column').addEventListener('change', updateExportButton);
  document.getElementById('export-pdf').addEventListener('click', exportToPDF);

  function updateExportButton() {
    const exportColumn = document.getElementById('export-column').value;
    const exportButton = document.getElementById('export-pdf');
    exportButton.disabled = !exportColumn;
  }

  function populateColumnSelects(sheetData) {
    // Populate export column select with top-level elements
    const exportColumnSelect = document.getElementById('export-column');
    exportColumnSelect.innerHTML = '<option value="">Select a section...</option>';
    
    // Get all top-level elements (children of root)
    if (sheetData.root && sheetData.root.children) {
      sheetData.root.children.forEach(child => {
        const option = document.createElement('option');
        option.value = child.value;
        option.textContent = child.value;
        exportColumnSelect.appendChild(option);
      });
    }
  }

  function renderProperties(node) {
    if (!node.properties || node.properties.length === 0) {
      return '';
    }

    return `
      <div class="property-row">
        ${node.properties.map(prop => `
          <div class="excel-node-property">
            <div class="property-name">${prop.columnName}</div>
            <div class="property-value">${prop.value || ''}</div>
          </div>
        `).join('')}
      </div>
    `;
  }

  function renderNodeForPDF(node) {
    let html = '';
    
    // Skip rendering for level 0 (Blok) nodes
    if (node.level === 0) {
      return '';
    }
    
    // For week nodes (level 1), render header and children
    if (node.level === 1) {
      html += `<div class="excel-node level-1" style="background-color: transparent !important;">`;
      html += `<div class="excel-node-header" style="background-color: transparent !important;">`;
      html += `<div class="node-title" style="font-size: 1rem; font-weight: bold; color: #333;">${node.value}</div>`;
      html += `</div>`;
      
      // Render children if they exist
      if (node.children && node.children.length > 0) {
        html += '<div class="excel-node-children" style="padding: 0; margin-left: 1rem;">';
        node.children.forEach(child => {
          html += renderNodeForPDF(child);
        });
        html += '</div>';
      }
      
      html += `</div>`;
    } else {
      // For Les nodes (level 2), render header, properties, and children
      html += `<div class="excel-node level-2" style="margin: 0.5rem 0; padding: 0.5rem; background-color: transparent !important;">`;
      html += `<div class="excel-node-header" style="background-color: rgba(74, 111, 165, 0.05) !important;">`;
      html += `<div class="node-title">${node.value}</div>`;
      html += `</div>`;
      
      // Add properties for Les nodes
      if (node.properties && node.properties.length > 0) {
        html += `<div class="excel-node-properties">`;
        html += `<div class="property-row" style="background-color: rgb(52 163 215 / 3%) !important;">`;
        node.properties.forEach(prop => {
          html += `
            <div class="excel-node-property" style="background-color: white !important;">
              <div class="property-name">${prop.columnName}</div>
              <div class="property-value">${prop.value || ''}</div>
            </div>
          `;
        });
        html += `</div>`;
        html += `</div>`;
      }
      
      // Render children if they exist
      if (node.children && node.children.length > 0) {
        html += '<div class="excel-node-children" style="padding: 0; margin-left: 1rem;">';
        node.children.forEach(child => {
          html += renderNodeForPDF(child);
        });
        html += '</div>';
      }
      
      html += `</div>`;
    }
    
    return html;
  }

  async function exportToPDF() {
    const selectedValue = document.getElementById('export-column').value;
    if (!selectedValue) {
      alert('Please select a section to export');
      return;
    }

    try {
      const exportButton = document.getElementById('export-pdf');
      exportButton.disabled = true;
      exportButton.textContent = 'Generating PDF...';

      // Get the active sheet ID
      const activeSheet = document.querySelector('.sheet-content.active');
      const sheetId = activeSheet.id.replace('sheet-', '');
      const sheetData = window.excelData.sheetsLoaded[sheetId];

      if (!sheetData) {
        throw new Error('No sheet data available');
      }

      // Find the selected top-level element
      const selectedNode = sheetData.root.children.find(node => node.value === selectedValue);
      if (!selectedNode) {
        throw new Error('Selected section not found');
      }

      // Create a temporary container for PDF content
      const tempContainer = document.createElement('div');
      tempContainer.className = 'pdf-container';

      // Add title
      const title = document.createElement('h1');
      title.className = 'pdf-title';
      title.textContent = `Materialenlijst - ${selectedValue}`;
      tempContainer.appendChild(title);

      // Create a section for the selected node
      const section = document.createElement('div');
      section.className = 'pdf-section';

      // Render all children recursively
      if (selectedNode.children) {
        selectedNode.children.forEach(child => {
          const childSection = document.createElement('div');
          childSection.className = 'pdf-section';
          // Add the PDF title to each section instead of blok-header
          const pdfTitle = document.createElement('div');
          pdfTitle.className = 'pdf-title';
          pdfTitle.textContent = selectedValue;
          childSection.appendChild(pdfTitle);
          childSection.innerHTML += renderNodeForPDF(child);
          section.appendChild(childSection);
        });
      }

      tempContainer.appendChild(section);

      // Generate PDF
      const response = await fetch('/pdf/generate', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          html: tempContainer.innerHTML,
          filename: `${selectedValue.replace(/[^a-z0-9]/gi, '_').toLowerCase()}.pdf`
        })
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error || 'Failed to generate PDF');
      }

      // Create a blob from the response
      const blob = await response.blob();
      
      // Create a download link and trigger it
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${selectedValue.replace(/[^a-z0-9]/gi, '_').toLowerCase()}.pdf`;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);

    } catch (error) {
      console.error('Error generating PDF:', error);
      alert(`Failed to generate PDF: ${error.message}`);
    } finally {
      const exportButton = document.getElementById('export-pdf');
      exportButton.disabled = false;
      exportButton.textContent = 'Export to PDF';
    }
  }
  
  // Add this function to show the preview
  function showPDFPreview(selectedValue, sheetData) {
    // Create modal container
    const modal = document.createElement('div');
    modal.className = 'pdf-preview-modal';
    
    // Create modal content
    const modalContent = document.createElement('div');
    modalContent.className = 'pdf-preview-content';
    
    // Create close button
    const closeButton = document.createElement('button');
    closeButton.className = 'pdf-preview-close';
    closeButton.innerHTML = '&times;';
    closeButton.onclick = () => modal.remove();
    
    // Create preview container
    const previewContainer = document.createElement('div');
    previewContainer.className = 'pdf-container';
    previewContainer.style.cssText = 'transform-origin: top left; transform: scale(1);';
    
    // Create content wrapper
    const contentWrapper = document.createElement('div');
    contentWrapper.className = 'pdf-content';
    
    // Add title
    const title = document.createElement('h1');
    title.className = 'pdf-title';
    title.style.cssText = 'font-size: 1.2rem; font-weight: bold; color: #34a3d7; background-color: transparent !important;';
    title.textContent = `Materialenlijst - ${selectedValue}`;
    contentWrapper.appendChild(title);
    
    // Find the selected top-level element
    const selectedNode = sheetData.root.children.find(node => node.value === selectedValue);
    if (!selectedNode) {
      throw new Error('Selected section not found');
    }

    // Create a section for the selected node
    const section = document.createElement('div');
    section.className = 'pdf-section';

    // Render all children recursively
    if (selectedNode.children) {
      selectedNode.children.forEach(child => {
        const childSection = document.createElement('div');
        childSection.className = 'pdf-section';
        // Add the PDF title to each section instead of blok-header
        const pdfTitle = document.createElement('div');
        pdfTitle.className = 'pdf-title';
        pdfTitle.style.cssText = 'font-size: 1.2rem; font-weight: bold; color: #34a3d7; background-color: transparent !important;';
        pdfTitle.textContent = selectedValue;
        childSection.appendChild(pdfTitle);
        childSection.innerHTML += renderNodeForPDF(child);
        section.appendChild(childSection);
      });
    }

    contentWrapper.appendChild(section);
    previewContainer.appendChild(contentWrapper);
    modalContent.appendChild(closeButton);
    modalContent.appendChild(previewContainer);
    modal.appendChild(modalContent);
    document.body.appendChild(modal);
  }

  // Update the export controls HTML to include a preview button
  function updateExportControls(sheetData) {
    const exportControls = document.querySelector('.export-controls');
    if (!exportControls) return;

    // Clear existing controls
    exportControls.innerHTML = '';

    // Create select element
    const select = document.createElement('select');
    select.id = 'export-column';
    select.innerHTML = '<option value="">Select a section to export</option>';
    
    // Add options for each top-level node
    if (sheetData.root && sheetData.root.children) {
      sheetData.root.children.forEach(node => {
        const option = document.createElement('option');
        option.value = node.value;
        option.textContent = node.value;
        select.appendChild(option);
      });
    }

    // Create preview button
    const previewButton = document.createElement('button');
    previewButton.id = 'preview-pdf';
    previewButton.className = 'control-button';
    previewButton.textContent = 'Preview PDF';
    previewButton.onclick = () => {
      const selectedValue = select.value;
      if (!selectedValue) {
        alert('Please select a section to preview');
        return;
      }
      showPDFPreview(selectedValue, sheetData);
    };

    // Create export button
    const exportButton = document.createElement('button');
    exportButton.id = 'export-pdf';
    exportButton.className = 'control-button';
    exportButton.textContent = 'Export to PDF';
    exportButton.onclick = exportToPDF;

    // Add elements to controls
    exportControls.appendChild(select);
    exportControls.appendChild(previewButton);
    exportControls.appendChild(exportButton);
  }
  
  initialize();
}); 