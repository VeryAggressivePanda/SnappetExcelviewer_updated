# Template Systeem - Excel Viewer

## Concept Overview

Het template systeem zorgt voor **uniforme structuren** met **contextuele data** uit Excel bestanden.

## Hoe het werkt

### 1. Template Parent (Eerste Node)
- De **eerste node** die wordt aangemaakt uit een kolom wordt de **template/master**
- Deze node heeft een speciale status: 🏗️ template indicator
- Alle aanpassingen aan deze node worden automatisch gekopieerd naar alle andere nodes

### 2. Duplicate Parents (Sibling Nodes)
- Alle **andere nodes** uit dezelfde kolom zijn **duplicates** van de template
- Ze volgen exact dezelfde structuur als de template
- Ze kunnen **niet individueel** worden aangepast

### 3. Structuur Replicatie
- **Structuurwijzigingen** aan de template worden automatisch toegepast op alle duplicates:
  - Child toevoegen → Alle duplicates krijgen dezelfde child
  - Layout wijzigen → Alle duplicates krijgen dezelfde layout
  - Kolom toewijzen → Alle duplicates krijgen dezelfde kolom-structuur

### 4. Contextuele Data Populatie
- **Data** wordt per duplicate **contextueel** ingevuld uit Excel:
  - Template (Blok 1) + Week child → Vult weeks uit Excel rijen waar Blok = "Blok 1"
  - Duplicate (Blok 2) + Week child → Vult weeks uit Excel rijen waar Blok = "Blok 2"

## Voorbeeld

```
Excel Data:
Blok 1, Week 1, Les A
Blok 1, Week 2, Les B  
Blok 2, Week 3, Les C
Blok 2, Week 4, Les D
```

**Template actie:** Bij Blok 1 (🏗️) voeg "Week" child toe

**Resultaat:**
- Blok 1 (Template): Week children → "Week 1", "Week 2"
- Blok 2 (Duplicate): Week children → "Week 3", "Week 4"

## Belangrijke Principes

1. **Eén template, alle duplicates volgen**
2. **Structuur is universeel, data is contextueel**
3. **Alleen template kan worden aangepast**
4. **Automatische replicatie naar alle siblings** 