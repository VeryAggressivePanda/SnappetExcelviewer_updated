# List View A4 Export Systeem - Excel Viewer

## Concept Overview

Het list view systeem zorgt voor **professionele A4 PDF exports** van Excel materialen met **automatische pagina-indeling** en **multi-course ondersteuning**.

## Hoe het werkt

### 1. Sheet Types & Layout Detection
- **Single Course**: Sheets zonder Course kolom → Horizontale layout (1:1:3 ratio)
- **Multi-Course**: Sheets met Course kolom → Verticale layout per course
- **Automatische detectie**: Systeem herkent type en past layout aan

### 2. A4 Pagina Rendering
- **Correcte afmetingen**: 210mm x 297mm A4 formaat
- **Echte hoogte meting**: Dynamische berekening van content per rij
- **Page breaks**: Intelligente verdeling over meerdere pagina's
- **Titel compensatie**: Minder content op pagina's met titels

### 3. Export Opties
- **Individual Courses**: Selecteer specifieke course voor export
- **Complete Overview**: Alle courses in één PDF
- **Dropdown populatie**: Automatisch gevuld met beschikbare courses

### 4. Multi-Course Features
- **Course Headers**: Tonen alleen groep (bijv. "Groep 4")
- **Herhalende Titels**: "Materialenlijst + Sheet naam" op eerste pagina elke course
- **Separate Paginering**: Elke course start op nieuwe pagina
- **Consistent Styling**: Uniforme opmaak across courses

## Layout Specificaties

### Single Course Layout
```
A4 Pagina (Horizontaal):
┌─────────────────────────────────────┐
│ Materialenlijst                     │
│ [Sheet naam]                        │
│                                     │
│ ┌─────┐ ┌─────┐ ┌─────────────────┐ │
│ │Col 1│ │Col 2│ │     Col 3       │ │
│ │(40%)│ │(20%)│ │     (40%)       │ │
│ └─────┘ └─────┘ └─────────────────┘ │
└─────────────────────────────────────┘
```

### Multi-Course Layout
```
A4 Pagina per Course (Verticaal):
┌─────────────────────────────────────┐
│ Materialenlijst                     │
│ [Sheet naam]                        │
│                                     │
│ [Groep 4] ← Course header           │
│                                     │
│ ┌─────────────────────────────────┐ │
│ │         Col 1                   │ │
│ ├─────────────────────────────────┤ │
│ │         Col 2                   │ │
│ ├─────────────────────────────────┤ │
│ │         Col 3                   │ │
│ └─────────────────────────────────┘ │
└─────────────────────────────────────┘
```

## Course Name Parsing

### Patronen Herkenning
- **Haakjes**: `"Snappet Rekenen (Groep 4)"` → `"Groep 4"`
- **Direct**: `"02. RekenWereld groep 3"` → `"groep 3"`
- **Case insensitive**: Werkt met "Groep" en "groep"
- **Fallback**: Volledige naam als geen patroon matcht

## Technische Details

### Page Break Berekening
1. **Content Meting**: Echte DOM rendering voor precisie
2. **Hoogte Compensatie**: ~1515px per A4 pagina (gemeten ratio)
3. **Titel Ruimte**: -80px voor pagina's met titels
4. **Row Synchronisatie**: Gelijke hoogte across kolommen

### Export Process
1. **Data Filtering**: Per course of alle courses
2. **DOM Generation**: A4 pagina's met correcte styling
3. **PDF Conversion**: html2canvas + jsPDF
4. **Download**: Automatische bestandsnaam generatie

## Belangrijke Principes

1. **Responsive Layout**: Past zich aan aan content type
2. **Professional Output**: Print-ready A4 formatting
3. **Efficient Pagination**: Optimale content verdeling
4. **Consistent Branding**: Uniforme titels en styling
5. **User-Friendly**: Intuïtieve export opties 