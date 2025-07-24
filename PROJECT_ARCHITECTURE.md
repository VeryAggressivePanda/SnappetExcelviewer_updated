# Project Architecture - Excel Viewer

## Project Overview

Een Excel Viewer applicatie met hiërarchische data visualisatie en template systeem voor uniforme structuren met contextuele data uit Excel bestanden.

## Core Functionaliteiten

### 1. Excel Bestandsverwerking
- Upload en parsing van .xlsx/.xls bestanden
- Automatische sheet detectie en metadata extractie
- Raw data extractie voor manual structure building
- Hiërarchische data organisatie

### 2. Template Systeem
- **Template Parent**: Eerste node uit kolom wordt master template (🏗️)
- **Duplicate Parents**: Andere nodes volgen template structuur
- **Structuur Replicatie**: Wijzigingen aan template worden gekopieerd
- **Contextuele Data**: Data wordt per duplicate contextueel ingevuld

### 3. Viewer Interface
- Interactieve hiërarchische weergave
- Expand/collapse functionaliteit
- Live editor voor structuur aanpassingen
- Customizable styling (fonts, themes, layouts)

### 4. PDF Export
- HTML naar PDF conversie
- Styled output met custom CSS
- Print-optimized layouts
- Preview functionaliteit

## Technische Stack

### Backend
- **Node.js** + **Express** - Server framework
- **ExcelJS** - Excel bestandsverwerking
- **Puppeteer** - PDF generatie
- **Multer** - File upload handling
- **EJS** - Template engine

### Frontend
- **Vanilla JavaScript** - Client-side interactiviteit
- **CSS3** - Styling en layouts
- **AbeZeh Font** - Custom typography

## Huidige Architectuur

```
Materialen_per_methode/
├── src/
│   ├── index.js              # Main server entry
│   ├── routes/
│   │   ├── excelRoutes.js    # Excel processing routes
│   │   └── pdfRoutes.js      # PDF generation routes
│   ├── views/
│   │   ├── index.ejs         # Upload page
│   │   └── viewer.ejs        # Main viewer interface
│   ├── public/
│   │   ├── css/styles.css    # All styling
│   │   └── js/
│   │       ├── viewer.js     # Main frontend logic (3500+ lines)
│   │       └── live-editor.js # Live editing functionality
│   ├── fonts/                # AbeZeh font files
│   ├── Excel/                # Sample Excel files
│   ├── uploads/              # Temporary file storage
│   └── temp/                 # PDF generation temp files
├── package.json
└── README.md
```

## Geplande Nieuwe Structuur

```
Materialen_per_methode/
├── src/
│   ├── server/
│   │   ├── index.js
│   │   ├── config/
│   │   ├── controllers/
│   │   ├── routes/
│   │   ├── services/
│   │   └── middleware/
│   ├── client/
│   │   ├── components/
│   │   ├── modules/
│   │   ├── utils/
│   │   └── styles/
│   ├── shared/
│   │   ├── types/
│   │   └── constants/
│   ├── assets/
│   │   ├── fonts/
│   │   └── samples/
│   └── views/
├── docs/
├── tests/
└── tools/
```

## Key Components Breakdown

### Frontend Modules (uit viewer.js)
1. **ExcelDataManager** - Data loading en caching
2. **HierarchyBuilder** - Structuur creatie en management  
3. **NodeRenderer** - UI rendering van nodes
4. **StructureEditor** - Live editing functionaliteit
5. **StyleManager** - Styling en theming
6. **ExportManager** - PDF preview en export
7. **EventManager** - Event handling en user interactions

### Backend Services
1. **ExcelService** - File processing en parsing
2. **PDFService** - PDF generatie
3. **FileService** - Upload en file management
4. **CacheService** - Data caching

## Data Flow

1. **Upload**: Excel file → ExcelService → Raw data extraction
2. **Processing**: Raw data → HierarchyBuilder → Structured nodes
3. **Rendering**: Structured nodes → NodeRenderer → HTML output
4. **Export**: HTML → PDFService → Styled PDF

## Template System Flow

1. **Template Creation**: Eerste node uit kolom wordt template
2. **Structure Definition**: Template krijgt child nodes en properties
3. **Replication**: Structuur wordt gekopieerd naar alle duplicates
4. **Data Population**: Contextuele data wordt ingevuld per duplicate

## Performance Considerations

- **Lazy Loading**: Sheet data wordt on-demand geladen
- **Caching**: Processed data wordt gecached in memory
- **Modular Loading**: Frontend modules worden modulair geladen
- **Optimized Rendering**: Efficient DOM updates en virtualization 