# PDM-SW v50.0 - Codifica gerarchica MACHINE/GROUP

## Data: 17 febbraio 2026

## 🎯 Obiettivo
Estensione del sistema di codifica per supportare versioni di macchine e gruppi con formato gerarchico a 3 livelli.

## 📋 Schema di codifica implementato

### Livello 1: MACHINE (Macchina)
- **Formato**: `MMM-V####`
- **Esempio**: `QQQ-V0001`, `QQQ-V0002`
- **Descrizione**: Versioni complete di macchina
- **Archiviazione**: `MACHINES/MMM/wip|rel|inrev|rev/`

### Livello 2: GROUP (Gruppo)
- **Formato**: `MMM_GGGG-V####`
- **Esempio**: `QQQ_1000-V0001`, `QQQ_1000-V0002`
- **Descrizione**: Versioni di gruppo macchina
- **Archiviazione**: `GROUPS/MMM/GGGG/wip|rel|inrev|rev/`

### Livello 3: PART/ASSY (invariato)
- **PART**: `MMM_GGGG-0001` → incrementale
- **ASSY**: `MMM_GGGG-9999` → decrementale
- **Con variante**: `MMM_GGGG-SKL-0001`, `MMM_GGGG-SKL-9999`
- **Archiviazione**: `MMM/GGGG/wip|rel|inrev|rev/` (come prima)

## 🔧 Modifiche implementate

### 1. Core Models
- **models.py**: Esteso `DocType` con `"MACHINE"` e `"GROUP"`

### 2. Configurazione
- **config.py**: Aggiunto segmento `"VNUM"` configurabile (default 4 cifre)
- **config.json**: Aggiornato con nuovo segmento per workspace esistenti

### 3. Generazione codici
- **codegen.py**: 
  - `build_machine_code()`: genera `MMM-V####`
  - `build_group_code()`: genera `MMM_GGGG-V####`
  - `build_code()`: invariato per PART/ASSY

### 4. Database
- **store.py**: 
  - Nuova tabella `ver_counters` per contatori MACHINE/GROUP
  - `allocate_ver_seq()`: alloca sequenze versioni
  - `allocate_seq()`: invariato per PART/ASSY

### 5. Archiviazione
- **archive.py**: 
  - `archive_dirs_for_machine()`: cartelle per MACHINE
  - `archive_dirs_for_group()`: cartelle per GROUP
  - `ext_for_doc_type()`: MACHINE/GROUP usano `.sldasm`

### 6. Interfaccia utente
- **app.py**: 
  - Nuovi pulsanti "Nuova MACCHINA" e "Nuovo GRUPPO" in tab Codifica
  - `_create_machine_version()`: crea versione macchina
  - `_create_group_version()`: crea versione gruppo
  - Gerarchia aggiornata con visualizzazione MACHINE/GROUP
  - Tag colorati: MACHINE (arancione), GROUP (blu)

### 7. Report
- **report_mixin.py**: 
  - Report gerarchico include MACHINE e GROUP
  - Conteggi separati per tipo documento

## 🎨 UI Features

### Tab Codifica
```
┌─────────────────────────────────────────────┐
│ VERSIONI (Macchine e Gruppi)               │
│ [Nuova MACCHINA (MMM-V####)]               │
│ [Nuovo GRUPPO (MMM_GGGG-V####)]            │
├─────────────────────────────────────────────┤
│ Tipo: [PART/ASSY] MMM: [...] GGGG: [...]  │
│ [Crea solo codice] [Crea + file SW]        │
└─────────────────────────────────────────────┘
```

### Tab Gerarchia
```
MMM: QQQ - MACCHINA Q
  📦 QQQ-V0001 [MACHINE] - Macchina versione 1
  GGGG: 1000 - GRUPPO 1000 (GROUP:2 PART:10 ASSY:3)
    📁 QQQ_1000-V0001 [GROUP] - Gruppo versione 1
    📁 QQQ_1000-V0002 [GROUP] - Gruppo versione 2
    QQQ_1000-0001 [PART]
    QQQ_1000-9999 [ASSY]
```

## 📊 Struttura archiviazione

```
C:/ArchivioCAD/
├── MACHINES/
│   └── QQQ/
│       ├── wip/
│       │   └── QQQ-V0001.sldasm
│       ├── rel/
│       ├── inrev/
│       └── rev/
├── GROUPS/
│   └── QQQ/
│       └── 1000/
│           ├── wip/
│           │   └── QQQ_1000-V0001.sldasm
│           ├── rel/
│           ├── inrev/
│           └── rev/
└── QQQ/
    └── 1000/
        ├── wip/
        │   ├── QQQ_1000-0001.sldprt
        │   ├── QQQ_1000-SKL-0001.sldprt
        │   └── QQQ_1000-9999.sldasm
        ├── rel/
        ├── inrev/
        └── rev/
```

## ⚙️ Configurazione

### Segmento VNUM
```json
{
  "code": {
    "segments": {
      "VNUM": {
        "enabled": true,
        "length": 4,
        "charset": "NUM",
        "case": "UPPER"
      }
    }
  }
}
```

### Personalizzazione
- `length`: numero di cifre (default 4 → V0001-V9999)
- Esempi:
  - `length: 3` → `QQQ-V001`
  - `length: 5` → `QQQ-V00001`

## 🔄 Migrazione dati

**Non richiesta**: i documenti PART/ASSY esistenti restano invariati.

### Compatibilità
- ✅ Workspace esistenti: compatibili al 100%
- ✅ Database: nuova tabella `ver_counters` creata automaticamente
- ✅ Config: segmento `VNUM` aggiunto automaticamente al primo avvio

## 📝 Note implementative

### Contatori
- **MACHINE**: contatore per MMM (gggg='')
- **GROUP**: contatore per (MMM, GGGG)
- **PART/ASSY**: contatori separati esistenti (next_part, next_assy)

### Template SolidWorks
- MACHINE usa `template_assembly`
- GROUP usa `template_assembly`
- PART usa `template_part`
- ASSY usa `template_assembly`

### Workflow
MACHINE e GROUP seguono lo stesso workflow di PART/ASSY:
- Stati: WIP → REL → IN_REV → OBS
- Revisioni gestite allo stesso modo
- Backup e log attività compatibili

## ✅ Verifiche effettuate

- [x] Nessun errore di compilazione
- [x] DocType esteso correttamente
- [x] Contatori versioni funzionanti
- [x] UI aggiornata con nuovi pulsanti
- [x] Gerarchia visualizza MACHINE/GROUP
- [x] Report include nuovi tipi
- [x] Config workspace aggiornato
- [x] Archiviazione cartelle separate

## 🚀 Prossimi passi suggeriti

1. Test creazione MACHINE/GROUP in ambiente reale
2. Verifica integrazione SolidWorks con nuovi tipi
3. Eventuale estensione macro runtime (se necessario)
4. Backup database prima del primo uso in produzione

---

**Versione app aggiornata**: v49.5 → v50.0
