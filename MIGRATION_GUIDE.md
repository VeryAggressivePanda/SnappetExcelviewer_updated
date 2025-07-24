# Migration Guide - Project Reorganization

## Overview

Het Excel Viewer project is georganiseerd van een monolithische structuur naar een modulaire architectuur voor betere onderhoudbaarheid, schaalbaarheid en ontwikkelervaring.

## Structural Changes

### Before (Oude Structuur)
```
Materialen_per_methode/
├── src/
│   ├── index.js              # Server entry point
│   ├── routes/
│   │   ├── excelRoutes.js    # Excel handling
│   │   └── pdfRoutes.js      # PDF generation
│   ├── views/
│   │   ├── index.ejs         # Upload page
│   │   └── viewer.ejs        # Main viewer
│   ├── public/
│   │   ├── css/styles.css    # All styling (1000+ lines)
│   │   └── js/
│   │       ├── viewer.js     # Main logic (3500+ lines)
│   │       └── live-editor.js
│   ├── fonts/                # Font assets
│   ├── Excel/                # Sample files
│   ├── uploads/              # File uploads
│   └── temp/                 # PDF temp files
├── package.json
└── README.md
```

### After (Nieuwe Structuur)
```
Materialen_per_methode/
├── src/
│   ├── server/               # Backend code
│   │   ├── index.js          # Server entry point
│   │   ├── config/           # Configuration files
│   │   ├── controllers/      # Business logic controllers
│   │   ├── routes/           # API routes
│   │   │   ├── excelRoutes.js
│   │   │   └── pdfRoutes.js
│   │   ├── services/         # Business services
│   │   └── middleware/       # Custom middleware
│   ├── client/               # Frontend code
│   │   ├── components/       # UI components
│   │   ├── modules/          # JavaScript modules
│   │   │   ├── viewer.js     # Main viewer logic
│   │   │   └── live-editor.js
│   │   ├── styles/           # CSS stylesheets
│   │   │   └── styles.css
│   │   └── utils/            # Client utilities
│   ├── shared/               # Shared code
│   │   ├── types/            # Type definitions
│   │   └── constants/        # Shared constants
│   ├── assets/               # Static assets
│   │   ├── fonts/            # Font files
│   │   └── samples/          # Sample Excel files
│   ├── views/                # EJS templates
│   ├── uploads/              # File uploads (unchanged)
│   └── temp/                 # PDF temp files (unchanged)
├── docs/                     # Documentation
├── tests/                    # Test files
├── tools/                    # Build and utility scripts
├── package.json              # Updated with new scripts
├── PROJECT_ARCHITECTURE.md   # Architecture documentation
└── MIGRATION_GUIDE.md        # This file
```

## Key Changes

### 1. Server Structure
- **Moved**: `src/index.js` → `src/server/index.js`
- **Updated**: Relative paths for views, static files, and uploads
- **Prepared**: Directory structure for controllers, services, and middleware

### 2. Client Structure
- **Moved**: `src/public/css/` → `src/client/styles/`
- **Moved**: `src/public/js/` → `src/client/modules/`
- **Updated**: Asset URLs in EJS templates
- **Prepared**: Directory structure for components and utilities

### 3. Assets Organization
- **Moved**: `src/fonts/` → `src/assets/fonts/`
- **Moved**: `src/Excel/` → `src/assets/samples/`
- **Organized**: Static assets in logical groupings

### 4. Package Configuration
- **Updated**: Main entry point to `src/server/index.js`
- **Enhanced**: npm scripts with build, test, lint, and clean commands
- **Added**: Keywords, license, and engine specifications

## Path Updates

### In Code Files

**Server (index.js):**
- Views path: `'views'` → `'../views'`
- Static path: `'public'` → `'../client'`
- Uploads path: `'uploads'` → `'../uploads'`

**Routes (excelRoutes.js, pdfRoutes.js):**
- Uploads path: `'../uploads'` → `'../../uploads'`
- Temp path: `'../temp'` → `'../../temp'`

**Templates (EJS files):**
- CSS: `/css/styles.css` → `/styles/styles.css`
- JS: `/js/viewer.js` → `/modules/viewer.js`

## Breaking Changes

### None - Backward Compatible
Deze reorganisatie behoudt alle functionaliteit. De applicatie werkt exact hetzelfde voor eindgebruikers.

### Development Changes
- Start script: blijft `npm start`
- Development: blijft `npm run dev`
- Entry point: automatisch bijgewerkt in package.json

## Next Steps (Planned)

### 1. Module Splitting
- Split `viewer.js` (3500+ lines) in logische modules:
  - `ExcelDataManager`
  - `HierarchyBuilder` 
  - `NodeRenderer`
  - `StructureEditor`
  - `StyleManager`
  - `ExportManager`
  - `EventManager`

### 2. Backend Services
- Extract business logic naar service classes
- Implement proper controllers
- Add configuration management
- Create middleware for common functionality

### 3. Shared Types
- Define TypeScript-style interfaces
- Create shared constants
- Add validation schemas

### 4. Testing & Tooling
- Add test framework setup
- Create build tools
- Add linting configuration
- Implement cleanup utilities

## Verification

To verify the migration was successful:

1. **Start the server**: `npm start`
2. **Check functionality**: Upload an Excel file
3. **Test features**: Verify viewer, editing, and PDF export work
4. **Check assets**: Ensure fonts and styling load correctly

The server should start without errors and all features should work exactly as before. 