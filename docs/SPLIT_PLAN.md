# Viewer.js Split Plan

## 🎯 Doel
Viewer.js (3582 regels) opsplitsen in 8 modules van max 300-450 regels elk.

## 📊 Huidige Analyse
- **45 functies** geïdentificeerd
- **Logische groepering** mogelijk op basis van functionaliteit
- **Geen breaking changes** - behoud alle huidige functionaliteit

## 🗂️ Module Opdeling

### 1. **core.js** (~300 regels)
**Lijnen:** 1-300  
**Functies:**
- DOM element initialization  
- Global data store setup
- populateExportColumnDropdown()
- saveExportSelection()

**Verantwoordelijkheid:** Basis setup en export dropdown management

### 2. **export.js** (~350 regels)  
**Lijnen:** 300-650
**Functies:**
- exportToPdf()
- previewPdf() 
- showExactPdfPreview()

**Verantwoordelijkheid:** PDF export en preview functionaliteit

### 3. **data-processing.js** (~400 regels)
**Lijnen:** 650-1050
**Functies:**
- showHierarchyConfigModal()
- processExcelData()
- processChildrenWithCurrentValues()
- cleanupNode()
- addPropertiesToNode()

**Verantwoordelijkheid:** Excel data verwerking en hierarchy building

### 4. **sheet-management.js** (~350 regels)
**Lijnen:** 1050-1400
**Functies:**
- loadSheetData()
- saveHierarchyConfiguration()
- saveCurrentHierarchyConfiguration()
- renderSheet()
- truncateMiddle()

**Verantwoordelijkheid:** Sheet loading, saving en basis rendering

### 5. **node-renderer.js** (~450 regels)
**Lijnen:** 1400-1850
**Functies:**
- renderNode()
- showAddChildElementModal()
- updatePreview() (eerste)
- applyStyles()

**Verantwoordelijkheid:** HTML rendering van nodes en styling

### 6. **ui-controls.js** (~350 regels)
**Lijnen:** 1850-2200
**Functies:**
- toggleAllNodes()
- initialize()
- getExcelCoordinates()
- getExcelCellValue()
- addElementDirectly()

**Verantwoordelijkheid:** UI controls en Excel coordinate management

### 7. **node-management.js** (~450 regels)
**Lijnen:** 2200-2650
**Functies:**
- getSavedHierarchyConfiguration()
- findParentNode()
- storeExpandState()
- addNode()
- deleteNode()
- showAddLesModal()

**Verantwoordelijkheid:** Node CRUD operaties en state management

### 8. **structure-builder.js** (~500 regels)
**Lijnen:** 2650-3582
**Functies:**
- addParentLevel()
- ensureLevelStyles()
- updateLevels()
- addChildContainer()
- addSiblingContainer()
- deleteContainer()
- initToolbar()
- showStructureBuilder()
- createStructureNode()
- extractStructure()
- applyCustomStructure()
- processStructureRow()

**Verantwoordelijkheid:** Advanced structure editing en toolbar

## 🔄 Migration Strategy

### Phase 1: Preparation
1. **Backup** origineel viewer.js
2. **Create module files** met basic structure
3. **Test current functionality** works

### Phase 2: Safe Extraction (1 module per stap)
1. Extract **core.js** → test
2. Extract **export.js** → test  
3. Extract **data-processing.js** → test
4. etc...

### Phase 3: Integration
1. **Update viewer.ejs** om alle modules te laden
2. **Test complete workflow**
3. **Remove original viewer.js**

## 🧪 Testing per Module
Voor elke module:
- [ ] Server start zonder errors
- [ ] Excel upload werkt
- [ ] Specifieke functionaliteit van module werkt
- [ ] PDF export werkt (eind-tot-eind test)

## 📁 File Structure After Split
```
src/client/modules/
├── viewer/ 
│   ├── core.js              # DOM setup, basic state
│   ├── export.js            # PDF export & preview  
│   ├── data-processing.js   # Excel data processing
│   ├── sheet-management.js  # Sheet loading & saving
│   ├── node-renderer.js     # HTML rendering
│   ├── ui-controls.js       # UI controls & coordinates
│   ├── node-management.js   # Node CRUD operations
│   └── structure-builder.js # Advanced structure editing
└── live-editor.js           # (unchanged)
```

## 🎯 Success Criteria
- ✅ All current functionality preserved
- ✅ Each module < 450 lines
- ✅ Clear separation of concerns  
- ✅ Easy to debug and maintain
- ✅ No performance degradation 