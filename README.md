# Excel Viewer

Web application voor het importeren en visualiseren van Excel bestanden met hiërarchische data weergave en template systeem.

## ✨ Features

### 📊 Excel Verwerking
- Import van Excel files (.xlsx/.xls) met drag & drop
- Automatische sheet detectie en metadata extractie
- Ondersteuning voor complexe hiërarchische data structuren
- Real-time data loading en caching

### 🏗️ Template Systeem
- **Template Parent**: Eerste node wordt master template (🏗️ indicator)
- **Duplicate Parents**: Automatische structuur replicatie naar sibling nodes
- **Contextuele Data**: Data wordt per duplicate contextueel ingevuld
- **Unified Structure**: Wijzigingen aan template worden gekopieerd naar alle duplicates

### 🎨 Interactieve Viewer
- Hiërarchische boom weergave met expand/collapse
- Live structuur editor voor real-time aanpassingen
- Customizable styling (fonts, themes, cell padding)
- Responsive layout voor verschillende schermformaten

### 📄 PDF Export
- High-quality PDF generatie met Puppeteer
- Styled output met custom CSS en typography
- Print-optimized layouts met AbeZeh font
- Preview functionaliteit voor WYSIWYG editing

## 🚀 Quick Start

### Prerequisites
- Node.js >= 16.0.0
- npm of yarn package manager

### Installation

1. **Clone het project**
   ```bash
   git clone <repository-url>
   cd Materialen_per_methode
   ```

2. **Installeer dependencies**
   ```bash
   npm install
   ```

3. **Start de applicatie**
   ```bash
   # Production mode
   npm start
   
   # Development mode met auto-reload
   npm run dev
   ```

4. **Open in browser**
   ```
   http://localhost:3001
   ```

## 📁 Project Structuur

```
Materialen_per_methode/
├── src/
│   ├── server/               # Backend (Node.js/Express)
│   │   ├── index.js          # Server entry point
│   │   ├── routes/           # API routes
│   │   ├── services/         # Business services (planned)
│   │   └── controllers/      # Logic controllers (planned)
│   ├── client/               # Frontend
│   │   ├── modules/          # JavaScript modules
│   │   ├── styles/           # CSS stylesheets
│   │   ├── components/       # UI components (planned)
│   │   └── utils/            # Client utilities (planned)
│   ├── assets/               # Static assets
│   │   ├── fonts/            # AbeZeh typography
│   │   └── samples/          # Sample Excel files
│   ├── shared/               # Shared code (planned)
│   └── views/                # EJS templates
├── docs/                     # Documentation
├── tools/                    # Build scripts (planned)
└── tests/                    # Test files (planned)
```

## 🔧 Usage

### 1. Excel Upload
1. Upload een Excel bestand via de web interface
2. Selecteer het gewenste sheet uit de tabs
3. De data wordt automatisch geladen en gerenderd

### 2. Template Systeem
1. **Template Creation**: De eerste node uit een kolom wordt automatisch template
2. **Structure Building**: Voeg child nodes toe aan de template
3. **Auto Replication**: Wijzigingen worden gekopieerd naar alle sibling nodes
4. **Data Population**: Data wordt contextueel ingevuld per duplicate

### 3. Customization
- **Styling**: Pas font size, theme, en cell padding aan
- **Layout**: Toggle grid view en expand/collapse states
- **Structure**: Live editing van hiërarchie via de structure builder

### 4. Export
1. Selecteer een kolom voor export via de dropdown
2. Preview de PDF output in de browser
3. Download de geformatteerde PDF

## 🛠️ Development

### Available Scripts

```bash
npm start        # Start production server
npm run dev      # Start development server met nodemon
npm run build    # Build project (placeholder)
npm test         # Run tests (placeholder)
npm run lint     # Lint code (placeholder)
npm run clean    # Clean temp files (placeholder)
```

### Key Technologies

- **Backend**: Node.js, Express, ExcelJS, Puppeteer
- **Frontend**: Vanilla JavaScript, CSS3, EJS templates
- **Typography**: AbeZeh font family
- **Processing**: JSON-based hierarchical data structures

## 📚 Documentation

- [Project Architecture](./PROJECT_ARCHITECTURE.md) - Complete technical overview
- [Migration Guide](./MIGRATION_GUIDE.md) - Structural changes documentation
- [Template System](./uitleg.md) - Template systeem uitleg

## 🎯 Roadmap

### Phase 1: Foundation ✅
- [x] Project reorganization
- [x] Modular directory structure
- [x] Updated documentation

### Phase 2: Modularization (Planned)
- [ ] Split viewer.js (3500+ lines) in modules
- [ ] Extract backend services
- [ ] Create reusable UI components
- [ ] Add shared type definitions

### Phase 3: Enhancement (Planned)
- [ ] Test framework implementatie
- [ ] Build pipeline automation
- [ ] Performance optimizations
- [ ] Extended customization options

## 🤝 Contributing

1. Fork het project
2. Create feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit changes (`git commit -m 'Add AmazingFeature'`)
4. Push naar branch (`git push origin feature/AmazingFeature`)
5. Open een Pull Request

## 📝 License

MIT License - zie LICENSE file voor details.

## 🆘 Support

Voor vragen of problemen:
1. Check de [documentatie](./docs/)
2. Zoek in bestaande [issues](issues)
3. Open een nieuwe issue met gedetailleerde beschrijving

---

**Excel Viewer** - Transforming spreadsheets into meaningful hierarchical visualizations. 