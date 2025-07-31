/**
 * Export Module - PDF Export & Preview Functionality
 * Using html2canvas + jsPDF for high-quality PDF generation
 */

// Function to handle PDF export using html2canvas + jsPDF
async function exportToPdf() {
  try {
    const activeSheetId = window.excelData.activeSheetId;
    if (!activeSheetId || !window.excelData.sheetsLoaded[activeSheetId]) {
      alert('Please select a sheet first');
      return;
    }
    
    // Get the sheet container to capture
    const sheetContainer = document.querySelector('.sheet-content.active .sheet-container');
    if (!sheetContainer) {
      alert('Could not find sheet content to export');
      return;
    }
    
    // Get the PDF button for loading state
    const { exportPdfButton } = window.ExcelViewerCore.getDOMElements();
    
    // Create title for PDF
    const sheetTitle = document.querySelector(`.tab-button[data-sheet-id="${activeSheetId}"]`).textContent;
    const pdfTitle = `${sheetTitle} - Export`;
    
    // Show loading state
    const originalText = exportPdfButton.textContent;
    exportPdfButton.textContent = 'Generating PDF...';
    exportPdfButton.disabled = true;
    
    try {
      // Wait a bit for libraries to be available
      let attempts = 0;
      while ((!window.html2canvas || !window.jsPDF) && attempts < 10) {
        await new Promise(resolve => setTimeout(resolve, 100));
        attempts++;
      }
      
      // Check if libraries are available
      if (typeof window.html2canvas !== 'function') {
        throw new Error('html2canvas is not available');
      }
      
      // jsPDF is available as window.jspdf.jsPDF in UMD builds
      if (!window.jsPDF && !window.jspdf) {
        throw new Error('jsPDF is not available');
      }
      
      console.log('Libraries loaded successfully');
      console.log('html2canvas:', typeof window.html2canvas);
      console.log('jsPDF:', typeof window.jsPDF);
      console.log('jspdf:', typeof window.jspdf);

      // Create PDF using jsPDF
      const jsPDF = window.jsPDF || window.jspdf?.jsPDF || window.jspdf;
      if (!jsPDF) {
        throw new Error('jsPDF constructor not found');
      }
      const pdf = new jsPDF('p', 'mm', 'a4');

      // Check if we have A4 pages created by the website
      let a4Pages = sheetContainer.querySelectorAll('.a4-page');
      
      // For multi-course: export selected course or complete overview
      const isMultiCourse = sheetContainer.hasAttribute('data-has-courses');
      if (isMultiCourse && a4Pages.length > 0) {
        const exportColumn = document.getElementById('export-column');
        const selectedValue = exportColumn?.value;
        
        if (!selectedValue) {
          alert('Please select a course to export');
          return;
        }
        
        if (selectedValue === 'Complete Overview') {
          // Export all courses - keep all A4 pages
          console.log(`Multi-course: exporting complete overview (${a4Pages.length} pages)`);
        } else if (selectedValue.startsWith('course:')) {
          // Export specific course
          const selectedCourse = selectedValue.replace('course:', '');
          a4Pages = sheetContainer.querySelectorAll(`.a4-page[data-course="${selectedCourse}"]`);
          console.log(`Multi-course: exporting "${selectedCourse}" (${a4Pages.length} pages)`);
        } else {
          alert('Please select a valid course to export');
          return;
        }
      }
      
      if (a4Pages.length > 0) {
        console.log(`Found ${a4Pages.length} A4 pages to export`);
        
        // Export each A4 page separately
        for (let i = 0; i < a4Pages.length; i++) {
          if (i > 0) {
            pdf.addPage();
          }
          
          console.log(`Capturing page ${i + 1}/${a4Pages.length}`);
          const canvas = await window.html2canvas(a4Pages[i], {
            useCORS: true,
            allowTaint: false,
            scale: 2, // Higher resolution for better quality
            backgroundColor: '#ffffff',
            logging: false,
            width: a4Pages[i].scrollWidth,
            height: a4Pages[i].scrollHeight
          });
          
          const imgData = canvas.toDataURL('image/jpeg', 0.95); // Higher quality
          const imgWidth = 210; // A4 width
          const imgHeight = (canvas.height * imgWidth) / canvas.width;
          
          // Add image to PDF page
          pdf.addImage(imgData, 'JPEG', 0, 0, imgWidth, Math.min(imgHeight, 297));
          console.log(`Added page ${i + 1} to PDF`);
        }
        
      } else {
        // Fallback: normal export for non-A4 pages
        console.log('No A4 pages found, using fallback export');
        
        // For multi-course: export selected course container
        let exportTarget = sheetContainer;
        if (isMultiCourse) {
          const exportColumn = document.getElementById('export-column');
          const selectedValue = exportColumn?.value;
          
          if (!selectedValue || !selectedValue.startsWith('course:')) {
            alert('Please select a course to export');
            return;
          }
          
          const selectedCourse = selectedValue.replace('course:', '');
          const selectedCourseContainer = sheetContainer.querySelector(`[data-course-container="${selectedCourse}"]`);
          if (selectedCourseContainer) {
            exportTarget = selectedCourseContainer;
            console.log(`Multi-course fallback: exporting "${selectedCourse}" container`);
          }
        }
        
        const canvas = await window.html2canvas(exportTarget, {
          useCORS: true,
          allowTaint: false,
          scale: 2, // Higher resolution for better quality
          backgroundColor: '#ffffff',
          logging: false,
          width: exportTarget.scrollWidth,
          height: exportTarget.scrollHeight
        });
        
        console.log('Canvas generated:', canvas.width + 'x' + canvas.height);
        
        // Convert to image data
        const imgData = canvas.toDataURL('image/jpeg', 0.95); // Higher quality
        
        // Calculate dimensions for A4
        const imgWidth = 210; // A4 width in mm
        const pageHeight = 297; // A4 height in mm
        const imgHeight = (canvas.height * imgWidth) / canvas.width;

        let heightLeft = imgHeight;
        let position = 0;

        // Add image to PDF with proper pagination
        pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight);
        heightLeft -= pageHeight;

        while (heightLeft >= 0 && position > -20000) { // Safety check
          position = heightLeft - imgHeight;
          pdf.addPage();
          pdf.addImage(imgData, 'JPEG', 0, position, imgWidth, imgHeight);
          heightLeft -= pageHeight;
        }
      }
      
      // Download the PDF
      pdf.save(`${pdfTitle}.pdf`);
      console.log('[PDF EXPORT] PDF generated and downloaded successfully');
      
    } catch (error) {
      console.error('[PDF EXPORT] Generation error:', error);
      throw error;
    }
    
    // Reset button state
    exportPdfButton.textContent = originalText;
    exportPdfButton.disabled = false;
    
  } catch (error) {
    console.error('[PDF EXPORT] Error:', error);
    alert(`Error generating PDF: ${error.message}`);
    
    // Reset button state
    const { exportPdfButton } = window.ExcelViewerCore.getDOMElements();
    if (exportPdfButton) {
      exportPdfButton.textContent = 'Export to PDF';
      exportPdfButton.disabled = false;
    }
  }
}

// Function to handle PDF preview
async function previewPdf() {
  try {
    const activeSheetId = window.excelData.activeSheetId;
    if (!activeSheetId || !window.excelData.sheetsLoaded[activeSheetId]) {
      alert('Please select a sheet first');
      return;
    }
    
    // Get the sheet container to preview
    const sheetContainer = document.querySelector('.sheet-content.active .sheet-container');
    if (!sheetContainer) {
      alert('Could not find sheet content to preview');
      return;
    }
    
    // Create title for preview
    const sheetTitle = document.querySelector(`.tab-button[data-sheet-id="${activeSheetId}"]`).textContent;
    const previewTitle = `${sheetTitle} - Preview`;
    
    // Wait for html2canvas to be available
    let attempts = 0;
    while (!window.html2canvas && attempts < 10) {
      await new Promise(resolve => setTimeout(resolve, 100));
      attempts++;
    }
    
    if (typeof window.html2canvas !== 'function') {
      throw new Error('html2canvas is not available for preview');
    }
    
    // Generate canvas from HTML for preview
    const canvas = await window.html2canvas(sheetContainer, {
      useCORS: true,
      allowTaint: false,
      scale: 1, // Consistent with export
      backgroundColor: '#ffffff',
      logging: false,
      width: sheetContainer.scrollWidth,
      height: sheetContainer.scrollHeight
    });
    
    // Show preview in modal
    showCanvasPreview(canvas, previewTitle);
    
  } catch (error) {
    console.error('[PDF PREVIEW] Error:', error);
    alert(`Error generating preview: ${error.message}`);
  }
}

// Function to show canvas preview in a modal
function showCanvasPreview(canvas, title) {
  // Create modal container
  const modal = document.createElement('div');
  modal.className = 'pdf-preview-modal';
  
  // Create content container
  const content = document.createElement('div');
  content.className = 'pdf-preview-content';
  
  // Add close button
  const closeButton = document.createElement('button');
  closeButton.className = 'pdf-preview-close';
  closeButton.innerHTML = '&times;';
  closeButton.addEventListener('click', () => {
    document.body.removeChild(modal);
  });
  
  // Add title
  const titleElement = document.createElement('h2');
  titleElement.textContent = 'PDF Preview: ' + title;
  titleElement.style.margin = '0 0 15px 0';
  titleElement.style.color = '#34a3d7';
  
  // Add buttons container
  const buttonsContainer = document.createElement('div');
  buttonsContainer.style.display = 'flex';
  buttonsContainer.style.justifyContent = 'space-between';
  buttonsContainer.style.margin = '10px 0';
  buttonsContainer.style.gap = '10px';
  
  // Add export button
  const exportButton = document.createElement('button');
  exportButton.className = 'control-button';
  exportButton.textContent = 'Export to PDF';
  exportButton.addEventListener('click', () => {
    exportToPdf();
    document.body.removeChild(modal);
  });
  
  buttonsContainer.appendChild(exportButton);
  
  // Add canvas to modal
  const canvasContainer = document.createElement('div');
  canvasContainer.style.width = '100%';
  canvasContainer.style.height = 'calc(100% - 100px)';
  canvasContainer.style.overflow = 'auto';
  canvasContainer.style.border = '1px solid #ddd';
  canvasContainer.style.backgroundColor = 'white';
  canvasContainer.style.padding = '20px';
  canvasContainer.style.textAlign = 'center';
  
  // Scale canvas to fit
  canvas.style.maxWidth = '100%';
  canvas.style.height = 'auto';
  canvas.style.boxShadow = '0 0 10px rgba(0,0,0,0.1)';
  
  canvasContainer.appendChild(canvas);
  
  // Assemble the modal
  content.appendChild(closeButton);
  content.appendChild(titleElement);
  content.appendChild(buttonsContainer);
  content.appendChild(canvasContainer);
  modal.appendChild(content);
  
  // Add to document
  document.body.appendChild(modal);
}

// Export functions for other modules
window.ExcelViewerExport = {
  exportToPdf,
  previewPdf,
  showCanvasPreview
}; 