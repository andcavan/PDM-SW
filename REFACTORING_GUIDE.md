# REFACTORING GUIDE - PDM-SW app.py

## Stato Corrente

**File:** `app.py`  
**Righe totali:** 4024  
**Problema:** File monolitico con tutti i tab embedati nella classe PDMApp

## Obiettivo

Modularizzare l'applicazione estraendo i tab in classi separate nel package `pdm_sw.gui/`.

---

## ✅ Completato

### 1. BaseTab (pdm_sw/gui/base_tab.py)
**Stato:** ✅ COMPLETATO

Classe base per tutti i tab che fornisce:
- Riferimenti comuni: `self.app`, `self.cfg`, `self.store`, `self.session`
- Metodi helper: `_validate_segment_strict()`, `_require_desc_upper()`, `_ask_large_text_input()`
- Funzioni utility: `warn()`, `info()`, `ask()`

**Pattern di utilizzo:**
```python
from pdm_sw.gui.base_tab import BaseTab

class MyTab(BaseTab):
    def __init__(self, parent_frame, app, cfg, store, session):
        super().__init__(app, cfg, store, session)
        self.parent = parent_frame
        self._build_ui()
    
    def _build_ui(self):
        # Costruisci l'interfaccia nel self.parent
        pass
```

---

### 2. TabGeneratore (pdm_sw/gui/tab_generatore.py)
**Stato:** ✅ COMPLETATO  
**Righe estratte:** 898-1158 (260 righe)

**Metodi estratti:**
- `_ui_generatore()` → `_build_ui()`
- `_on_machine_list_selected()`
- `_selected_mmm()`
- `_selected_gggg()`
- `refresh_machines()` (era `_refresh_machines`)
- `refresh_groups()` (era `_refresh_groups`)
- `_on_group_machine_selected()`
- `_add_machine()`
- `_edit_machine_desc()`
- `_del_machine()`
- `_add_group()`
- `_edit_group_desc()`
- `_del_group()`

**Metodi pubblici che l'app deve chiamare:**
- `refresh_machines()` - richiamato da `_refresh_all()`
- `refresh_groups()` - richiamato da `_refresh_all()`

**Dipendenze app:**
- `self.app._refresh_machine_menus()` - aggiorna menu in altri tab
- `self.app._refresh_group_menu()` - aggiorna menu in altri tab
- `self.app._refresh_hierarchy_tree()` - aggiorna albero gerarchia

---

### 3. TabCodifica (pdm_sw/gui/tab_codifica.py)
**Stato:** ✅ COMPLETATO  
**Righe estratte:** 1426-1957 (531 righe)

**Metodi estratti:**
- `_ui_codifica()` → `_build_ui()`
- `refresh_machine_menus()` (era `_refresh_machine_menus`)
- `_refresh_group_menu()`
- `refresh_vvv_menu()` (era `_refresh_vvv_menu`)
- `_on_doc_type_change()`
- `_refresh_preview()`
- `_show_next_code()`
- `_browse_link_file()`
- `_clear_link_file()`
- `_import_linked_files_to_wip()`
- `_generate_document()`

**Metodi pubblici che l'app deve chiamare:**
- `refresh_machine_menus()` - richiamato da `_refresh_all()`
- `refresh_vvv_menu()` - richiamato da `_refresh_all()`

**Dipendenze app:**
- `self.app._create_files_for_code()` - crea file SW (troppo complesso per estrarre ora)
- `self.app._log_activity()` - logging attività
- `self.app._refresh_all()` - refresh completo dopo creazione documento

---

## 📋 Da Completare

### 4. TabSolidWorks (pdm_sw/gui/tab_solidworks.py)
**Stato:** ⏳ DA FARE  
**Righe da estrarre:** 1159-1425 (266 righe)

**Punto di partenza:** Linea 1159
```python
def _ui_solidworks(self):
```

**Metodi da estrarre:**
- `_ui_solidworks()` → `_build_ui()`
- `_test_solidworks()`
- `_publish_sw_macro()`
- `_browse_archive_root()`
- `_browse_template_part()`
- `_browse_template_assembly()`
- `_browse_template_drawing()`
- Gestori configurazione proprietà SW

**Metodi pubblici attesi:**
- Nessuno (tab configurazione, non refreshato da altri)

**Dipendenze app:**
- Import macro publish utilities
- Config manager per salvare settings

**Complessità:** 🟡 MEDIA (gestisce configurazioni e dialoghi)

---

### 5. TabGerarchia (pdm_sw/gui/tab_gerarchia.py)
**Stato:** ⏳ DA FARE  
**Righe da estrarre:** 1958-2694 (736 righe)

**Punto di partenza:** Linea 1958
```python
def _ui_gerarchia(self):
```

**Metodi da estrarre:**
- `_ui_gerarchia()` → `_build_ui()`
- `_refresh_hierarchy_tree()`
- `_on_hier_tree_select()`
- `_show_hier_doc_details()`
- Gestori visualizzazione dettagli documenti in gerarchia

**Metodi pubblici attesi:**
- `refresh_tree()` (era `_refresh_hierarchy_tree`)

**Dipendenze app:**
- `self.store.list_machines()`, `list_groups()`, `list_documents_in_group()`
- Widget Treeview complesso con icone

**Complessità:** 🟡 MEDIA (albero gerarchico con logica visualizzazione)

---

### 6. TabOperativo (pdm_sw/gui/tab_operativo.py)
**Stato:** ⏳ DA FARE  
**Righe da estrarre:** 2695-3481 (786 righe)

**Punto di partenza:** Linea 2695
```python
def _ui_operativo(self):
```

**Metodi da estrarre:**
- `_ui_operativo()` → `_build_ui()`
- `_refresh_rc_table()` - tabella ricerca/codifica
- `_refresh_workflow_panel()` - pannello workflow
- `_on_rc_table_select()` - selezione documento
- `_open_model()`, `_open_drw()` - apertura file SW
- `_edit_doc_description()` - modifica descrizione
- `_release_wip()`, `_create_inrev()`, `_approve_inrev()`, `_cancel_inrev()` - workflow
- `_set_obsolete()`, `_restore_obsolete()` - gestione obsoleti
- `_sync_sw_to_pdm()` - sincronizzazione proprietà SW
- Logica pannello split Ricerca/Workflow
- Lock/unlock documenti
- Gestione filtri e ricerche

**Metodi pubblici attesi:**
- `refresh_table()` (era `_refresh_rc_table`)
- `refresh_workflow()` (era `_refresh_workflow_panel`)

**Dipendenze app:**
- `self.store` - query complesse sui documenti
- SolidWorks API per apertura file
- `pdm_sw.archive` - operazioni workflow
- RCCopyMixin per copia dati da tabella

**Complessità:** 🔴 ALTA (tab più complesso, logica workflow, lock, split panel)

---

### 7. TabMonitor (pdm_sw/gui/tab_monitor.py)
**Stato:** ⏳ DA FARE  
**Righe da estrarre:** 3482-3987 (505 righe)

**Punto di partenza:** Linea 3482
```python
def _ui_monitor(self):
```

**Metodi da estrarre:**
- `_ui_monitor()` → `_build_ui()`
- `_refresh_monitor_panel()` - aggiorna tabella attività
- `_start_monitor_polling()` - polling automatico
- `_stop_monitor_polling()` - stop polling
- `_on_monitor_select()` - selezione attività
- `_show_activity_details()` - mostra dettagli JSON
- Gestione filtri temporali

**Metodi pubblici attesi:**
- `refresh()` (era `_refresh_monitor_panel`)
- `start_polling()`, `stop_polling()`

**Dipendenze app:**
- `self.store.list_activities()` - query log attività
- Timer per auto-refresh
- JSON viewer per dettagli

**Complessità:** 🟡 MEDIA (tabella log con auto-refresh)

---

### 8. TabGestioneCodifica (pdm_sw/gui/tab_gestione_codifica.py)
**Stato:** ⏳ DA FARE  
**Righe da estrarre:** ~600-897 (circa 297 righe, da verificare)

**Punto di partenza:** Da cercare `_ui_gestione_codifica`

**Metodi da estrarre:**
- `_ui_gestione_codifica()` → `_build_ui()`
- Gestori configurazione segmenti (MMM, GGGG, SEQ, VVV, VNUM)
- Gestori preset varianti
- Salvataggio configurazione

**Metodi pubblici attesi:**
- Nessuno (tab configurazione)

**Dipendenze app:**
- `self.cfg_mgr.save()` - salvataggio config
- Validatori segmenti

**Complessità:** 🟢 BASSA (form configurazione)

---

### 9. TabManuale (pdm_sw/gui/tab_manuale.py)
**Stato:** ⏳ DA FARE  
**Righe da estrarre:** Da identificare (probabilmente < 100 righe)

**Punto di partenza:** Da cercare `_ui_manuale`

**Metodi da estrarre:**
- `_ui_manuale()` → `_build_ui()`
- Testo statico con istruzioni uso

**Metodi pubblici attesi:**
- Nessuno (tab statico)

**Dipendenze app:**
- Nessuna

**Complessità:** 🟢 BASSA (solo testo)

---

## 🔧 Pattern di Integrazione in app.py

### Modifica __init__() della classe PDMApp

**Prima:**
```python
def __init__(self):
    # ... init esistente ...
    self._build_ui()
```

**Dopo:**
```python
def __init__(self):
    # ... init esistente ...
    self._build_ui()
    
    # Inizializza tab modulari
    self.tab_generatore_obj = None
    self.tab_codifica_obj = None
    # ... altri tab ...
```

### Modifica _build_ui()

**Prima:**
```python
def _build_ui(self):
    # ... creazione tabs ...
    self._ui_operativo()
    self._ui_gerarchia()
    self._ui_monitor()
    self._ui_codifica()
    self._ui_gestione_codifica()
    self._ui_generatore()
    self._ui_solidworks()
    self._ui_manuale()
```

**Dopo:**
```python
def _build_ui(self):
    # ... creazione tabs ...
    from pdm_sw.gui.tab_generatore import TabGeneratore
    from pdm_sw.gui.tab_codifica import TabCodifica
    # ... altri import ...
    
    self._ui_operativo()  # TODO: estrarre
    self._ui_gerarchia()  # TODO: estrarre
    self._ui_monitor()    # TODO: estrarre
    
    # Tab modulari
    self.tab_codifica_obj = TabCodifica(self.tab_cod, self, self.cfg, self.store, self.session)
    # self._ui_gestione_codifica()  # TODO: estrarre
    self.tab_generatore_obj = TabGeneratore(self.tab_gen, self, self.cfg, self.store, self.session)
    # self._ui_solidworks()  # TODO: estrarre
    # self._ui_manuale()  # TODO: estrarre
```

### Modifica _refresh_all()

**Prima:**
```python
def _refresh_all(self):
    self._set_ws_label()
    self._set_shared_root_label()
    self._refresh_machines()
    self._refresh_groups()
    self._refresh_machine_menus()
    self._refresh_vvv_menu()
    self._refresh_hierarchy_tree()
    self._refresh_rc_table()
    self._refresh_workflow_panel()
    self._refresh_monitor_panel()
```

**Dopo:**
```python
def _refresh_all(self):
    self._set_ws_label()
    self._set_shared_root_label()
    
    # Refresh tab modulari
    if self.tab_generatore_obj:
        self.tab_generatore_obj.refresh_machines()
        self.tab_generatore_obj.refresh_groups()
    if self.tab_codifica_obj:
        self.tab_codifica_obj.refresh_machine_menus()
        self.tab_codifica_obj.refresh_vvv_menu()
    
    # Refresh tab ancora da estrarre
    self._refresh_hierarchy_tree()  # TODO: tab_gerarchia_obj.refresh_tree()
    self._refresh_rc_table()        # TODO: tab_operativo_obj.refresh_table()
    self._refresh_workflow_panel()  # TODO: tab_operativo_obj.refresh_workflow()
    self._refresh_monitor_panel()   # TODO: tab_monitor_obj.refresh()
```

---

## 📊 Statistiche Refactoring

| Componente | Righe | Stato | Complessità |
|------------|-------|-------|-------------|
| **BaseTab** | ~190 | ✅ FATTO | 🟢 |
| **TabGeneratore** | 260 | ✅ FATTO | 🟢 |
| **TabCodifica** | 531 | ✅ FATTO | 🟡 |
| **TabSolidWorks** | 266 | ⏳ TODO | 🟡 |
| **TabGestioneCodifica** | ~297 | ⏳ TODO | 🟢 |
| **TabManuale** | <100 | ⏳ TODO | 🟢 |
| **TabGerarchia** | 736 | ⏳ TODO | 🟡 |
| **TabMonitor** | 505 | ⏳ TODO | 🟡 |
| **TabOperativo** | 786 | ⏳ TODO | 🔴 |
| **TOTALE** | ~3571 | **27% completato** | - |

**Righe rimanenti in app.py dopo refactoring completo:** ~450 righe  
(solo bootstrap, helpers comuni, _refresh_all, lifecycle)

---

## 🎯 Priorità Implementazione

1. ✅ **BaseTab** - Fondamenta
2. ✅ **TabGeneratore** - Funzionalità base, bassa complessità
3. ✅ **TabCodifica** - Core business logic
4. ⏳ **TabGestioneCodifica** - Configurazione (bassa complessità)
5. ⏳ **TabManuale** - Statico (triviale)
6. ⏳ **TabSolidWorks** - Configurazione SW
7. ⏳ **TabGerarchia** - Visualizzazione
8. ⏳ **TabMonitor** - Logging/monitoring
9. ⏳ **TabOperativo** - Più complesso, da affrontare per ultimo

---

## ⚠️ Note Importanti

### Metodi che rimangono in PDMApp (troppo accoppiati o condivisi)

- `_create_files_for_code()` - interazione complessa con SW API
- `_log_activity()` - logging centralizzato
- `_build_sw_props_for_doc()` - mapping proprietà SW
- `_sync_sw_to_pdm()` - sincronizzazione bidirezionale SW↔PDM
- `_norm_segment()` - normalizzazione segmenti
- Gestione workspace e config (lifecycle)
- Backup e session management

### Convenzioni Naming

- Metodi pubblici nei tab: **senza underscore** (`refresh_machines`)
- Metodi privati nei tab: **con underscore** (`_selected_mmm`)
- Metodi nell'app che i tab chiamano: **con underscore** (`self.app._refresh_all()`)

### Testing

Dopo ogni estrazione:
1. Verificare che app.py non abbia errori di sintassi
2. Avviare l'applicazione e verificare il tab estratto
3. Testare le funzionalità principali del tab
4. Verificare che `_refresh_all()` funzioni correttamente

---

## 📝 Workflow Estrazione Tab

1. **Identificare linee esatte** del metodo `_ui_NOMTAB()`
2. **Mappare metodi helper** associati (cerca `_` prefix + nome correlato)
3. **Creare classe** `TabNOME` che estende `BaseTab`
4. **Copiare metodi** mantenendo logica invariata
5. **Sostituire `self` con `self.app`** dove necessario (accesso a metodi app)
6. **Rinominare** metodo `_ui_NOME()` → `_build_ui()`
7. **Esporre metodi pubblici** rimuovendo `_` prefix dove appropriato
8. **Aggiornare app.py** import + inizializzazione + `_refresh_all()`
9. **Testare** funzionalità del tab
10. **Commit** con messaggio descrittivo

---

## 🚀 Prossimi Passi

- [ ] Estrarre TabGestioneCodifica (facile)
- [ ] Estrarre TabManuale (facile)
- [ ] Estrarre TabSolidWorks (medio)
- [ ] Estrarre TabGerarchia (medio)
- [ ] Estrarre TabMonitor (medio)
- [ ] Estrarre TabOperativo (difficile)
- [ ] Refactoring macro_runtime.py (1292 righe, separato)
- [ ] Documentazione API finale
