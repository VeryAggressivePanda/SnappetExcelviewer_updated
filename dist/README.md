# Excel Viewer - Static Build

This is a static build of the Excel Viewer application that can be deployed to any web hosting service.

## Features

- Client-side Excel file processing (no server required)
- PDF export functionality
- Custom ABeZeh font family integrated
- Hierarchical view and list template modes
- Supports .xlsx and .xls files
- Drag & drop file upload

## Deployment Instructions

1. Upload all contents of this folder to your web hosting service
2. Make sure index.html is accessible from your domain root
3. Ensure all assets are properly linked (check browser console for any 404 errors)

## Font Integration

The application uses the ABeZeh font family in all weights and styles:
- Thin (100) / ThinItalic
- ExtraLight (200) / ExtraLightItalic  
- Regular (400) / Italic
- Medium (500) / MediumItalic
- Bold (700) / BoldItalic
- ExtraBold (800) / ExtraBoldItalic

## Browser Requirements

- Modern browser with JavaScript enabled
- File API support for local file reading
- Canvas API for PDF generation
- Font loading support

## File Structure

- index.html - Main application file with embedded fonts
- assets/styles/ - CSS stylesheets
- assets/modules/ - JavaScript modules
- assets/fonts/ - ABeZeh font files (EOT, WOFF, WOFF2)
