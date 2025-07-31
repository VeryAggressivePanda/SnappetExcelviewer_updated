const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const router = express.Router();

// Set up multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(__dirname, '../../uploads');
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    // ALWAYS create new file with timestamp - no more reusing existing files
    cb(null, Date.now() + '-' + file.originalname);
  }
});

const upload = multer({ 
  storage,
  fileFilter: (req, file, cb) => {
    // Log incoming file info for debugging
    console.log('Uploaded file:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      extension: path.extname(file.originalname).toLowerCase()
    });
    
    // Accept Excel files - expanded list of mime types
    const validMimeTypes = [
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/excel',
      'application/x-excel',
      'application/x-msexcel'
    ];
    
    const validExtensions = ['.xlsx', '.xls'];
    const fileExtension = path.extname(file.originalname).toLowerCase();
    
    if (validExtensions.includes(fileExtension)) {
      // If it has a valid extension, accept it regardless of mimetype
      // Some browsers/systems might not report the correct mimetype
      return cb(null, true);
    } else {
      cb('Error: Only .xlsx or .xls files are allowed');
    }
  }
});

// Home page route
router.get('/', (req, res) => {
  res.render('index', { title: 'Excel Viewer' });
});

// In-memory cache for processed Excel files
const excelCache = new Map();

// Handle file upload
router.post('/upload', upload.single('excelFile'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).render('index', { 
        title: 'Excel Viewer',
        error: 'Please upload an Excel file'
      });
    }

    const filePath = req.file.path;
    const fileId = req.file.filename; // Use the stored filename directly
    console.log('Processing file:', filePath);
    console.log('File ID for cache and reference:', fileId);
    
    // Process the Excel file
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    // Extract metadata about sheets for tabs - only visible and non-empty sheets
    const sheetsMeta = [];
    const sheetsData = {};
    
    workbook.eachSheet((worksheet, sheetId) => {
      // Skip hidden sheets
      if (worksheet.state === 'hidden' || worksheet.state === 'veryHidden') {
        console.log(`Skipping hidden sheet: ${worksheet.name}`);
        return;
      }
      
      // Check if sheet is empty
      let isEmpty = true;
      let rowCount = 0;
      
      worksheet.eachRow(() => {
        isEmpty = false;
        rowCount++;
        return false; // Stop after finding one row
      });
      
      // Skip empty sheets or sheets with just one empty row
      if (isEmpty || (rowCount === 1 && worksheet.getRow(1).values.filter(Boolean).length === 0)) {
        console.log(`Skipping empty sheet: ${worksheet.name}`);
        return;
      }
      
      sheetsMeta.push({
        id: sheetId,
        name: worksheet.name
      });
      
      // DISABLED: No auto-processing during upload - let user build manually
      // if (sheetsMeta.length === 1) {
      //   sheetsData[sheetId] = processWorksheet(worksheet);
      // }
    });
    
    // If no valid sheets were found
    if (sheetsMeta.length === 0) {
      return res.status(400).render('index', {
        title: 'Excel Viewer',
        error: 'The Excel file does not contain any visible non-empty sheets'
      });
    }
    
    // Store in cache with metadata and first sheet data
    excelCache.set(fileId, {
      workbook,
      filePath,
      originalName: req.file.originalname,
      sheetsMeta,
      processedSheets: sheetsData
    });
    
    // Render the viewer page with the sheet metadata but not the data
    // The data will be loaded via AJAX
    res.render('viewer', { 
      title: 'Excel Viewer',
      filename: req.file.originalname,
      fileId: fileId,
      sheets: sheetsMeta
    });
    
  } catch (error) {
    console.error('Error processing Excel file:', error);
    res.status(500).render('index', {
      title: 'Excel Viewer',
      error: 'Error processing the Excel file: ' + error.message
    });
  }
});

// API endpoint to load sheet data on demand
router.get('/api/sheet/:fileId/:sheetId', (req, res) => {
  try {
    const { fileId, sheetId } = req.params;
    const sheetIdNum = parseInt(sheetId, 10);
    
    console.log(`Loading sheet data for file ${fileId}, sheet ${sheetIdNum}`);
    
    // Check if file exists in cache
    if (!excelCache.has(fileId)) {
      console.log(`File ${fileId} not found in cache`);
      return res.status(404).json({ error: 'File not found' });
    }
    
    const fileData = excelCache.get(fileId);
    console.log(`File found in cache, loaded data for ${fileData.originalName}`);
    
    // Always process fresh data for manual structure building
    // (Don't return cached hierarchical data - we want raw data for manual setup)
    
    // Get the worksheet
    const worksheet = fileData.workbook.getWorksheet(sheetIdNum);
    if (!worksheet) {
      console.log(`Worksheet ${sheetIdNum} not found in workbook`);
      return res.status(404).json({ error: 'Sheet not found' });
    }
    
    console.log(`Processing worksheet ${worksheet.name} (ID: ${sheetIdNum})`);
    
    // Check if the sheet is hidden
    if (worksheet.state === 'hidden' || worksheet.state === 'veryHidden') {
      console.log(`Sheet ${sheetIdNum} is hidden, skipping`);
      return res.status(403).json({ error: 'Sheet is hidden' });
    }
    
    // Check if sheet is empty
    let isEmpty = true;
    worksheet.eachRow(() => {
      isEmpty = false;
      return false; // Stop after finding one row
    });
    
    if (isEmpty) {
      console.log(`Sheet ${sheetIdNum} is empty`);
      return res.json({ 
        data: {
          headers: [],
          root: { type: 'root', children: [], level: -1 }
        }
      });
    }
    
    // Extract raw data for processing
    const rawData = [];
    let maxColumns = 0;
    
    // Determine column count
    if (worksheet.actualRowCount && worksheet.actualColumnCount) {
      maxColumns = worksheet.actualColumnCount;
      console.log(`Using actual column count: ${maxColumns}`);
    } else {
      worksheet.eachRow((row) => {
        maxColumns = Math.max(maxColumns, row.cellCount);
      });
      console.log(`Calculated max column count: ${maxColumns}`);
    }
    
    console.log(`Final column count: ${maxColumns}`);
    
    // Extract data
    worksheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
      const rowData = [];
      
      for (let i = 1; i <= maxColumns; i++) {
        const cell = row.getCell(i);
        let value = '';
        
        if (cell && cell.value !== null && cell.value !== undefined) {
          if (typeof cell.value === 'object') {
            if (cell.value instanceof Date) {
              value = cell.value.toLocaleDateString();
            } else if (cell.value.formula) {
              value = cell.value.result || '';
            } else if (cell.value.text) {
              value = cell.value.text;
            } else if (cell.value.richText) {
              value = cell.value.richText.map(rt => rt.text).join('');
            }
          } else {
            value = String(cell.value);
          }
        }
        
        rowData.push(value.trim());
      }
      
      rawData.push(rowData);
      if (rowIndex <= 2) {
        console.log(`Row ${rowIndex} data:`, rowData.slice(0, 5));
      }
    });
    
    console.log(`Extracted ${rawData.length} rows of raw data`);
    
    // ONLY return raw data for manual hierarchy building
    // Do NOT build automatic hierarchy that interferes with template system
    const headers = rawData.length > 0 ? rawData[0] : [];
    
    // Return ONLY raw data for manual structure building
    const response = {
      headers: headers,
      data: rawData, // Raw data for manual structure building
      isRawData: true, // Flag indicating this needs manual configuration
      needsConfiguration: true, // Flag for manual setup
      root: { 
        type: 'root', 
        children: [], 
        level: -1 
      } // Empty root structure for manual building
    };
    
    // Store in cache (without processing)
    console.log(`Storing raw data for sheet ${sheetIdNum} in cache`);
    fileData.processedSheets[sheetIdNum] = response;
    
    // Return the raw data for manual building
    console.log(`Sending raw data response for sheet ${sheetIdNum}`);
    return res.json({ data: response });
    
  } catch (error) {
    console.error('Error loading sheet data:', error);
    res.status(500).json({ error: 'Error loading sheet data: ' + error.message });
  }
});

// Function to process a worksheet into a hierarchical JSON structure
function processWorksheet(worksheet) {
  console.log(`Processing worksheet: ${worksheet.name}`);
  
  // First pass: read data as flat rows
  const rawRows = [];
  let maxColumns = 0;
  
  // First, determine the actual used range to get a more accurate max column count
  if (worksheet.actualRowCount && worksheet.actualColumnCount) {
    maxColumns = worksheet.actualColumnCount;
    console.log(`Using actual column count: ${maxColumns}`);
  } else {
    // Fallback if actualColumnCount is not available
    worksheet.eachRow((row) => {
      maxColumns = Math.max(maxColumns, row.cellCount);
    });
    console.log(`Calculated max column count: ${maxColumns}`);
  }
  
  // Remove the forced minimum of 10 columns
  console.log(`Final column count: ${maxColumns}`);
  
  // Map out the data including all rows (with empty cells)
  worksheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
    const rowData = [];
    rowData.rowIndex = rowIndex;
    
    // Process all possible columns to ensure grid alignment
    for (let i = 1; i <= maxColumns; i++) {
      const cell = row.getCell(i);
      let value = '';
      
      // Extract cell value with proper handling of different types
      if (cell && cell.value !== null && cell.value !== undefined) {
        if (typeof cell.value === 'object') {
          if (cell.value instanceof Date) {
            value = cell.value.toLocaleDateString();
          } else if (cell.value.formula) {
            value = cell.value.result || '';
          } else if (cell.value.text) {
            value = cell.value.text;
          } else if (cell.value.richText) {
            value = cell.value.richText.map(rt => rt.text).join('');
          }
        } else {
          value = String(cell.value);
        }
      }
      
      rowData.push(value.trim());
    }
    
    // Only log first few rows to avoid too much output
    if (rowIndex <= 3) {
      console.log(`Row ${rowIndex}:`, rowData.slice(0, 5));
    }
    
    rawRows.push(rowData);
  });
  
  // If no data, return empty array
  if (rawRows.length === 0) {
    console.warn('No data found in worksheet');
    return {
      headers: [],
      root: { type: 'root', children: [], level: -1 }
    };
  }
  
  // Identify headers (first row)
  const headers = rawRows[0] || [];
  
  // Log summary of data for debugging
  console.log(`Processed worksheet with ${rawRows.length} rows and ${maxColumns} columns`);
  console.log('First few headers:', headers.slice(0, 5));
  
  // Check if we have data rows other than the header
  if (rawRows.length <= 1) {
    console.warn('No data rows found (only header row)');
    return {
      headers: headers,
      root: { type: 'root', children: [], level: -1 }
    };
  }
  
  // DISABLED: No auto-hierarchy building - return raw data for manual structure building
  const rawDataResponse = {
    headers: headers,
    data: rawRows, // Raw data for manual structure building
    isRawData: true,
    needsConfiguration: true,
    root: { 
      type: 'root', 
      children: [], 
      level: -1 
    } // Empty root structure for manual building
  };
  
  console.log(`Raw data prepared with ${headers.length} headers and ${rawRows.length} rows`);
  
  return rawDataResponse;
}

// Function to build a hierarchical structure from flat data
function buildHierarchy(rows, headers) {
  console.log(`Building hierarchy from ${rows.length} rows`);
  
  // Root node to hold all top-level items
  const root = {
    type: 'root',
    children: [],
    level: -1
  };
  
  // Skip empty or invalid headers
  if (!headers || headers.length === 0) {
    console.warn('No headers available for hierarchy');
    return root;
  }
  
  // Find the hierarchy columns and their names
  const hierarchyColumns = {
    course: { index: -1, name: '' },
    blok: { index: -1, name: '' },
    week: { index: -1, name: '' },
    les: { index: -1, name: '' }
  };
  
  // Map headers to hierarchy columns
  headers.forEach((header, index) => {
    if (!header) return;
    
    const headerText = String(header).toLowerCase();
    if (headerText.includes('course')) {
      hierarchyColumns.course = { index, name: header };
    } else if (headerText.includes('blok')) {
      hierarchyColumns.blok = { index, name: header };
    } else if (headerText.includes('week')) {
      hierarchyColumns.week = { index, name: header };
    } else if (headerText.includes('les')) {
      hierarchyColumns.les = { index, name: header };
    }
  });
  
  // Log found hierarchy columns
  console.log('Found hierarchy columns:', hierarchyColumns);
  
  // Track current nodes at each level
  const currentNodes = {
    course: null,
    blok: null,
    week: null,
    les: null
  };
  
  // Process each row
  rows.forEach((row, rowIndex) => {
    // Skip completely empty rows
    if (row.every(cell => cell === '')) {
      return;
    }
    
    // Process course if present
    if (hierarchyColumns.course.index !== -1 && row[hierarchyColumns.course.index]) {
      const courseValue = row[hierarchyColumns.course.index].trim();
      if (courseValue) {
        let courseNode = root.children.find(child => child.value === courseValue);
        if (!courseNode) {
          courseNode = {
            type: 'item',
            value: courseValue,
            columnName: hierarchyColumns.course.name,
            level: 0,
            children: [],
            properties: []
          };
          root.children.push(courseNode);
        }
        currentNodes.course = courseNode;
      }
    }
    
    // Process blok if present
    if (hierarchyColumns.blok.index !== -1 && row[hierarchyColumns.blok.index]) {
      const blokValue = row[hierarchyColumns.blok.index].trim();
      if (blokValue) {
        const parentNode = currentNodes.course || root;
        let blokNode = parentNode.children.find(child => 
          child.value === blokValue && 
          child.columnName === hierarchyColumns.blok.name
        );
        if (!blokNode) {
          blokNode = {
            type: 'item',
            value: blokValue,
            columnName: hierarchyColumns.blok.name, // Always keep this as "Blok"
            level: 1, // Always level 1, regardless of parent
            children: [],
            properties: []
          };
          parentNode.children.push(blokNode);
        }
        currentNodes.blok = blokNode;
      }
    }
    
    // Process week if present
    if (hierarchyColumns.week.index !== -1 && row[hierarchyColumns.week.index]) {
      const weekValue = row[hierarchyColumns.week.index].trim();
      if (weekValue) {
        const parentNode = currentNodes.blok || currentNodes.course || root;
        let weekNode = parentNode.children.find(child => 
          child.value === weekValue && 
          child.columnName === hierarchyColumns.week.name
        );
        if (!weekNode) {
          weekNode = {
            type: 'item',
            value: weekValue,
            columnName: hierarchyColumns.week.name,
            level: 2,
            children: [],
            properties: []
          };
          parentNode.children.push(weekNode);
        }
        currentNodes.week = weekNode;
      }
    }
    
    // Process les if present
    if (hierarchyColumns.les.index !== -1 && row[hierarchyColumns.les.index]) {
      const lesValue = row[hierarchyColumns.les.index].trim();
      if (lesValue) {
        const parentNode = currentNodes.week || currentNodes.blok || currentNodes.course || root;
        let lesNode = parentNode.children.find(child => 
          child.value === lesValue && 
          child.columnName === hierarchyColumns.les.name
        );
        if (!lesNode) {
          lesNode = {
            type: 'item',
            value: lesValue,
            columnName: hierarchyColumns.les.name,
            level: 3,
            children: [],
            properties: []
          };
          parentNode.children.push(lesNode);
        }
        currentNodes.les = lesNode;
        
        // Add all columns after Les as CHILDREN, not properties
        // Start from the column after Les
        for (let i = hierarchyColumns.les.index + 1; i < headers.length; i++) {
          if (headers[i]) { // Only add if there's a header for this column
            const contentValue = row[i] || '';
            
            // Create a child node for each content column
            let contentChild = lesNode.children.find(child => 
              child.columnName === headers[i]
            );
            
            if (!contentChild) {
              contentChild = {
                type: 'item',
                value: contentValue,
                columnName: headers[i],
                level: 4, // Level 4 since Les is level 3
                children: [],
                properties: []
              };
              lesNode.children.push(contentChild);
            }
            
            // Also add as property for backwards compatibility
            lesNode.properties.push({
              column: i,
              columnName: headers[i],
              value: contentValue
            });
          }
        }
      }
    }
  });
  
  console.log(`Processed hierarchy with ${root.children.length} top-level items`);
  if (root.children.length > 0) {
    console.log('First course:', JSON.stringify(root.children[0]).substring(0, 200) + '...');
  }
  
  return root;
}

// Function to remove circular references for JSON serialization
function removeCircularReferences(data) {
  if (!data) {
    console.warn('removeCircularReferences called with null or undefined data');
    return {
      headers: [],
      root: { type: 'root', children: [], level: -1 }
    };
  }
  
  // For the headers and root structure
  if (data.headers && data.root) {
    // Create a deep copy
    return {
      headers: Array.isArray(data.headers) ? [...data.headers] : [],
      root: deepCopy(data.root)
    };
  }
  
  // If data doesn't have the expected structure, return a valid empty structure
  console.warn('Data missing expected structure in removeCircularReferences');
  return {
    headers: [],
    root: { type: 'root', children: [], level: -1 }
  };
}

// Function to perform a deep copy of an object, eliminating circular references
function deepCopy(obj) {
  // Handle null or undefined
  if (!obj) return null;
  
  // Handle primitive types
  if (typeof obj !== 'object') return obj;
  
  // Handle arrays
  if (Array.isArray(obj)) {
    return obj.map(item => deepCopy(item));
  }
  
  // Handle objects
  const copy = {};
  
  // Copy all properties except 'parent' to break circular references
  for (const key in obj) {
    if (key !== 'parent') {
      copy[key] = deepCopy(obj[key]);
    }
  }
  
  return copy;
}

module.exports = router; 