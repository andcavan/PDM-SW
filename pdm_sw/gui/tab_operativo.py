"""
Tab Operativo - Pannello unificato: Ricerca&Consultazione + Workflow.

Layout: paned window orizzontale con:
- Sinistra: tabella ricerca/consultazione + filtri
- Destra: workflow panel con info documento e transizioni stato

Salvato: ratio larghezza pannelli (workflow_width_ratio in local_settings.json)
"""
import tkinter as tk
from tkinter import ttk
from pathlib import Path
from typing import TYPE_CHECKING

import customtkinter as ctk

from .base_tab import BaseTab
from pdm_sw.ui.table import Table

if TYPE_CHECKING:
    from pdm_sw.models import Document

# Limiti ratio workflow width (0.2 = 20% minimo, 0.6 = 60% massimo)
WORKFLOW_WIDTH_RATIO_MIN = 0.2
WORKFLOW_WIDTH_RATIO_MAX = 0.6


class TabOperativo(BaseTab):
    """
    Tab Operativo: pannello unificato con ricerca/consultazione e workflow.
    
    Features:
    - Paned window con ratio salvabile
    - Ricerca documenti con filtri multipli
    - Workflow panel con transizioni stato
    - Delegazione business logic all'app principale
    """
    
    def __init__(self, parent_frame, app, cfg, store, session):
        """
        Inizializza il tab Operativo.
        
        Args:
            parent_frame: frame genitore in cui costruire la UI
            app: riferimento all'app principale
            cfg: configurazione workspace
            store: store SQLite
            session: sessione utente
        """
        super().__init__(app, cfg, store, session)
        self.root = parent_frame
        
        # Paned window
        self._operativo_paned = None
        
        # Ricerca&Consultazione (left panel)
        self.tab_rc = None
        self.search_text_var = None
        self.search_state_var = None
        self.search_type_var = None
        self.search_mmm_var = None
        self.search_gggg_var = None
        self.search_vvv_var = None
        self.include_obs_var = None
        self.rc_table = None
        
        # Workflow (right panel)
        self.tab_wf = None
        self._workflow_compact = True
        self.wf_code_var = None
        self.wf_state_var = None
        self.btn_send_to_wf = None
        self.wf_btn_release = None
        self.wf_btn_create_rev = None
        self.wf_btn_approve = None
        self.wf_btn_cancel = None
        self.wf_btn_obsolete = None
        self.wf_btn_restore_obs = None
        self.wf_info = None
        
        self._build_ui()
    
    def _build_ui(self):
        """Costruisce il layout paned con ricerca+workflow."""
        tab = self.root
        
        # Paned window orizzontale
        paned = ttk.Panedwindow(tab, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=8, pady=8)
        self._operativo_paned = paned
        
        # Left: Ricerca&Consultazione
        left = ctk.CTkFrame(paned)
        # Right: Workflow
        right = ctk.CTkFrame(paned)
        
        paned.add(left, weight=6)
        paned.add(right, weight=4)
        
        # Bind sash movement
        paned.bind("<B1-Motion>", lambda _e: self._on_operativo_sash_move(save=False))
        paned.bind("<ButtonRelease-1>", lambda _e: self._on_operativo_sash_move(save=True))
        
        # Costruisco i due pannelli
        self.tab_rc = left
        self.tab_wf = right
        self._build_ricerca_consultazione()
        self._build_workflow_panel()
        
        # Applico split ratio salvato dopo render
        self.app.after(60, self._apply_operativo_split)
    
    def _build_ricerca_consultazione(self):
        """Costruisce pannello sinistro: filtri + tabella ricerca."""
        tab = self.tab_rc
        
        # Barra azioni (gialla) con pulsanti operativi
        actions = ctk.CTkFrame(tab, fg_color="#F2D65C")
        actions.pack(fill="x", padx=10, pady=(10, 6))
        
        ctk.CTkButton(
            actions,
            text="AGGIORNA",
            width=120,
            command=self.refresh_table
        ).pack(side="left", padx=6, pady=6)
        
        ctk.CTkButton(
            actions,
            text="APRI MODELLO",
            width=140,
            command=lambda: self.app._open_selected_model()
        ).pack(side="left", padx=6, pady=6)
        
        ctk.CTkButton(
            actions,
            text="APRI DRW",
            width=120,
            command=lambda: self.app._open_selected_drw()
        ).pack(side="left", padx=6, pady=6)
        
        ctk.CTkButton(
            actions,
            text="CREA FILE MANCANTI",
            width=170,
            command=lambda: self.app._create_missing_files_for_selected()
        ).pack(side="left", padx=6, pady=6)
        
        ctk.CTkButton(
            actions,
            text="COPIA CODICE",
            width=130,
            command=lambda: self.app._copy_selected_code_to_new_wip()
        ).pack(side="left", padx=6, pady=6)
        
        ctk.CTkButton(
            actions,
            text="FORZA PDM->SW",
            width=150,
            command=lambda: self.app._force_pdm_to_sw_selected()
        ).pack(side="left", padx=6, pady=6)
        
        ctk.CTkButton(
            actions,
            text="FORZA SW->PDM",
            width=150,
            command=lambda: self.app._force_sw_to_pdm_selected()
        ).pack(side="left", padx=6, pady=6)
        
        # Barra filtri
        filters = ctk.CTkFrame(tab)
        filters.pack(fill="x", padx=10, pady=(0, 8))
        
        # Testo libero
        ctk.CTkLabel(filters, text="Testo:").pack(side="left", padx=(6, 2), pady=6)
        self.search_text_var = ctk.StringVar(value="")
        ctk.CTkEntry(filters, textvariable=self.search_text_var, width=220).pack(side="left", padx=4, pady=6)
        
        # Stato
        ctk.CTkLabel(filters, text="Stato:").pack(side="left", padx=(10, 2), pady=6)
        self.search_state_var = ctk.StringVar(value="")
        ctk.CTkComboBox(
            filters,
            variable=self.search_state_var,
            values=["", "WIP", "REL", "IN_REV", "OBS"],
            width=110
        ).pack(side="left", padx=4, pady=6)
        
        # Tipo
        ctk.CTkLabel(filters, text="Tipo:").pack(side="left", padx=(10, 2), pady=6)
        self.search_type_var = ctk.StringVar(value="")
        ctk.CTkComboBox(
            filters,
            variable=self.search_type_var,
            values=["", "PART", "ASSY"],
            width=110
        ).pack(side="left", padx=4, pady=6)
        
        # MMM
        ctk.CTkLabel(filters, text="MMM:").pack(side="left", padx=(10, 2), pady=6)
        self.search_mmm_var = ctk.StringVar(value="")
        ctk.CTkEntry(filters, textvariable=self.search_mmm_var, width=70).pack(side="left", padx=4, pady=6)
        
        # GGGG
        ctk.CTkLabel(filters, text="GGGG:").pack(side="left", padx=(10, 2), pady=6)
        self.search_gggg_var = ctk.StringVar(value="")
        ctk.CTkEntry(filters, textvariable=self.search_gggg_var, width=80).pack(side="left", padx=4, pady=6)
        
        # VVV
        ctk.CTkLabel(filters, text="VVV:").pack(side="left", padx=(10, 2), pady=6)
        self.search_vvv_var = ctk.StringVar(value="")
        ctk.CTkEntry(filters, textvariable=self.search_vvv_var, width=80).pack(side="left", padx=4, pady=6)
        
        # Mostra OBS
        self.include_obs_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            filters,
            text="Mostra OBS",
            variable=self.include_obs_var
        ).pack(side="left", padx=(14, 6), pady=6)
        
        # Pulsanti filtri
        ctk.CTkButton(
            filters,
            text="CERCA",
            width=120,
            command=self._search_rc
        ).pack(side="left", padx=(8, 6), pady=6)
        
        ctk.CTkButton(
            filters,
            text="RESET",
            width=110,
            command=self._reset_rc_filters
        ).pack(side="left", padx=6, pady=6)
        
        # Tabella risultati (schema costruito dinamicamente con proprietà SW)
        columns, headings, props, key_index = self.app._build_table_schema_with_sw_props()
        self.rc_table = Table(
            tab,
            columns=columns,
            headings=headings,
            on_double_click=lambda code: self.app._doc_dbl(code),
            key_index=key_index
        )
        self.rc_table.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # Bind selection per abilitare pulsante "INVIA A WORKFLOW"
        self.rc_table.tree.bind("<<TreeviewSelect>>", self._on_rc_select)
        
        # Prima ricerca automatica
        self.refresh_table()
    
    def _build_workflow_panel(self):
        """Costruisce pannello destro: workflow con transizioni stato."""
        compact = self._workflow_compact
        frame = ctk.CTkFrame(self.tab_wf)
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Top bar: codice selezionato + stato
        top = ctk.CTkFrame(frame)
        top.pack(fill="x", pady=(0, 8))
        
        self.wf_code_var = tk.StringVar(value="")
        self.wf_state_var = tk.StringVar(value="")
        
        if compact:
            ctk.CTkLabel(
                top,
                text="Workflow",
                font=ctk.CTkFont(size=15, weight="bold")
            ).pack(anchor="w", padx=6, pady=(6, 2))
            
            ctk.CTkLabel(top, text="Codice selezionato:").pack(anchor="w", padx=6, pady=(0, 2))
            ctk.CTkEntry(top, textvariable=self.wf_code_var).pack(fill="x", padx=6, pady=(0, 6))
            ctk.CTkLabel(top, textvariable=self.wf_state_var).pack(anchor="w", padx=6, pady=(0, 6))
            
            top_btns = ctk.CTkFrame(top, fg_color="transparent")
            top_btns.pack(fill="x", padx=4, pady=(0, 4))
            
            self.btn_send_to_wf = ctk.CTkButton(
                top_btns,
                text="INVIA A WORKFLOW",
                width=150,
                state="disabled",
                command=self._send_rc_to_workflow
            )
            self.btn_send_to_wf.pack(side="left", padx=4)
            
            ctk.CTkButton(
                top_btns,
                text="Carica",
                command=self.refresh_workflow,
                width=90
            ).pack(side="left", padx=4)
            
            ctk.CTkButton(
                top_btns,
                text="Report",
                command=lambda: self.app._generate_code_report(),
                width=90
            ).pack(side="left", padx=4)
            
            ctk.CTkLabel(
                top_btns,
                text="Ridimensiona trascinando il separatore centrale",
                text_color="#666666"
            ).pack(side="left", padx=8)
        else:
            ctk.CTkLabel(top, text="Codice selezionato:").pack(side="left", padx=6)
            ctk.CTkEntry(top, textvariable=self.wf_code_var, width=260).pack(side="left", padx=6)
            ctk.CTkLabel(top, textvariable=self.wf_state_var).pack(side="left", padx=10)
            
            self.btn_send_to_wf = ctk.CTkButton(
                top,
                text="INVIA A WORKFLOW",
                width=170,
                state="disabled",
                command=self._send_rc_to_workflow
            )
            self.btn_send_to_wf.pack(side="left", padx=6)
            
            ctk.CTkLabel(
                top,
                text="(ridimensiona trascinando il separatore)",
                text_color="#666666"
            ).pack(side="left", padx=(2, 8))
            
            ctk.CTkButton(
                top,
                text="Carica",
                command=self.refresh_workflow,
                width=120
            ).pack(side="left", padx=6)
            
            ctk.CTkButton(
                top,
                text="REPORT CODICE",
                command=lambda: self.app._generate_code_report(),
                width=150
            ).pack(side="left", padx=6)
        
        # Pulsanti workflow (transizioni stato)
        row1 = ctk.CTkFrame(frame, fg_color="transparent")
        row1.pack(fill="x", pady=(0, 4))
        row2 = ctk.CTkFrame(frame, fg_color="transparent")
        row2.pack(fill="x", pady=(0, 6))
        
        if compact:
            self.wf_btn_release = ctk.CTkButton(
                row1,
                text="WIP -> REL",
                command=lambda: self.app._wf_release()
            )
            self.wf_btn_create_rev = ctk.CTkButton(
                row1,
                text="REL -> IN_REV",
                command=lambda: self.app._wf_create_rev()
            )
            self.wf_btn_approve = ctk.CTkButton(
                row1,
                text="IN_REV -> REL (OK)",
                command=lambda: self.app._wf_approve()
            )
            self.wf_btn_cancel = ctk.CTkButton(
                row2,
                text="IN_REV -> REL (ANNULLA)",
                command=lambda: self.app._wf_cancel()
            )
            self.wf_btn_obsolete = ctk.CTkButton(
                row2,
                text="-> OBS",
                command=lambda: self.app._wf_obsolete()
            )
            self.wf_btn_restore_obs = ctk.CTkButton(
                row2,
                text="RIPRISTINA OBS",
                command=lambda: self.app._wf_restore_obs()
            )
        else:
            self.wf_btn_release = ctk.CTkButton(
                row1,
                text="WIP -> REL (Release)",
                command=lambda: self.app._wf_release()
            )
            self.wf_btn_create_rev = ctk.CTkButton(
                row1,
                text="REL -> IN_REV (Crea revisione)",
                command=lambda: self.app._wf_create_rev()
            )
            self.wf_btn_approve = ctk.CTkButton(
                row1,
                text="IN_REV -> REL (Approva)",
                command=lambda: self.app._wf_approve()
            )
            self.wf_btn_cancel = ctk.CTkButton(
                row2,
                text="IN_REV -> REL (Annulla)",
                command=lambda: self.app._wf_cancel()
            )
            self.wf_btn_obsolete = ctk.CTkButton(
                row2,
                text="-> OBS (Obsoleto)",
                command=lambda: self.app._wf_obsolete()
            )
            self.wf_btn_restore_obs = ctk.CTkButton(
                row2,
                text="OBS -> Ripristina",
                command=lambda: self.app._wf_restore_obs()
            )
        
        # Layout pulsanti
        if compact:
            row1.grid_columnconfigure(0, weight=1)
            row1.grid_columnconfigure(1, weight=1)
            row2.grid_columnconfigure(0, weight=1)
            row2.grid_columnconfigure(1, weight=1)
            
            self.wf_btn_release.grid(row=0, column=0, sticky="ew", padx=4, pady=3)
            self.wf_btn_create_rev.grid(row=0, column=1, sticky="ew", padx=4, pady=3)
            self.wf_btn_approve.grid(row=1, column=0, columnspan=2, sticky="ew", padx=4, pady=3)
            
            self.wf_btn_cancel.grid(row=0, column=0, columnspan=2, sticky="ew", padx=4, pady=3)
            self.wf_btn_obsolete.grid(row=1, column=0, sticky="ew", padx=4, pady=3)
            self.wf_btn_restore_obs.grid(row=1, column=1, sticky="ew", padx=4, pady=3)
        else:
            for b in (self.wf_btn_release, self.wf_btn_create_rev, self.wf_btn_approve):
                b.pack(side="left", padx=4, pady=4)
            for b in (self.wf_btn_cancel, self.wf_btn_obsolete, self.wf_btn_restore_obs):
                b.pack(side="left", padx=4, pady=4)
        
        # Info textbox
        self.wf_info = ctk.CTkTextbox(frame, height=(320 if compact else 240))
        self.wf_info.pack(fill="both", expand=True, pady=8)
        
        # Inizializza stato
        self._on_rc_select(None)
        self._update_workflow_buttons(None)
    
    # ========== PANED WINDOW MANAGEMENT ==========
    
    def _apply_operativo_split(self) -> None:
        """Applica lo split ratio salvato al paned window."""
        paned = self._operativo_paned
        if paned is None:
            return
        
        ratio = self._clamp_workflow_width_ratio(self.app.workflow_width_ratio)
        self.app.workflow_width_ratio = ratio
        
        try:
            total = int(paned.winfo_width() or 0)
            if total <= 20:
                return
            
            left_px = int(round((1.0 - ratio) * total))
            min_left = int(total * (1.0 - WORKFLOW_WIDTH_RATIO_MAX))
            max_left = int(total * (1.0 - WORKFLOW_WIDTH_RATIO_MIN))
            left_px = max(min_left, min(max_left, left_px))
            
            paned.sashpos(0, left_px)
        except Exception:
            pass
    
    def _on_operativo_sash_move(self, save: bool = False) -> None:
        """Gestisce il movimento del sash (separatore paned window)."""
        paned = self._operativo_paned
        if paned is None:
            return
        
        try:
            total = int(paned.winfo_width() or 0)
            if total <= 20:
                return
            
            sash_x = int(paned.sashpos(0))
            right_ratio = 1.0 - (float(sash_x) / float(total))
            self.app.workflow_width_ratio = self._clamp_workflow_width_ratio(right_ratio)
            
            if save:
                self.app._save_local_settings()
        except Exception:
            pass
    
    def _clamp_workflow_width_ratio(self, value: float) -> float:
        """Limita il ratio workflow width ai limiti permessi."""
        return max(WORKFLOW_WIDTH_RATIO_MIN, min(WORKFLOW_WIDTH_RATIO_MAX, value))
    
    # ========== RICERCA & CONSULTAZIONE (LEFT PANEL) ==========
    
    def _on_rc_select(self, _evt=None):
        """Gestisce selezione nella tabella ricerca (abilita pulsante workflow)."""
        code = self._get_selected_rc_code()
        if self.btn_send_to_wf is not None:
            self.btn_send_to_wf.configure(state="normal" if code else "disabled")
    
    def _get_selected_rc_code(self) -> str:
        """Ritorna il codice selezionato nella tabella ricerca."""
        try:
            sel = self.rc_table.tree.selection()
            if not sel:
                return ""
            values = self.rc_table.tree.item(sel[0], "values")
            # schema: M,D,CODICE,... -> code index 2
            return str(values[2]) if len(values) > 2 else ""
        except Exception:
            return ""
    
    def _send_rc_to_workflow(self):
        """Invia codice selezionato dalla ricerca al pannello workflow."""
        code = self._get_selected_rc_code()
        if not code:
            return
        
        self.wf_code_var.set(code)
        try:
            self.app.tabs.set("Operativo")
        except Exception:
            pass
        
        self.refresh_workflow()
    
    def _reset_rc_filters(self):
        """Resetta tutti i filtri di ricerca."""
        self.search_text_var.set("")
        self.search_state_var.set("")
        self.search_type_var.set("")
        self.search_mmm_var.set("")
        self.search_gggg_var.set("")
        self.search_vvv_var.set("")
        self.include_obs_var.set(False)
        self.refresh_table()
    
    def _search_rc(self):
        """Esegue ricerca con filtri correnti."""
        self.refresh_table()
    
    def refresh_table(self):
        """
        Aggiorna tabella ricerca/consultazione con filtri correnti.
        
        Chiamato da:
        - Pulsante AGGIORNA
        - CERCA / RESET filtri
        - _switch_workspace
        - _refresh_all
        """
        # Schema dinamico con proprietà SW
        columns, headings, props, key_index = self.app._build_table_schema_with_sw_props()
        self.rc_table.set_schema(columns=columns, headings=headings, key_index=key_index)
        
        # Leggi filtri
        txt = (self.search_text_var.get() or "").strip()
        st = (self.search_state_var.get() or "").strip()
        tp = (self.search_type_var.get() or "").strip()
        mmm = (self.search_mmm_var.get() or "").strip().upper()
        gggg = (self.search_gggg_var.get() or "").strip().upper()
        vvv = (self.search_vvv_var.get() or "").strip().upper()
        include_obs = bool(self.include_obs_var.get())
        
        # Query store
        docs = self.store.search_documents(
            text=txt,
            state=st if st else None,
            doc_type=tp if tp else None,
            mmm=mmm if mmm else None,
            gggg=gggg if gggg else None,
            vvv=vvv if vvv else None,
            include_obs=include_obs,
        )
        
        # Bulk load SW properties
        sw_values = self.store.get_custom_values_bulk([d.code for d in docs], props) if props else {}
        
        # Build rows
        rows = []
        for d in docs:
            m_ok, d_ok = self.app._model_and_drawing_flags(d)
            
            base = [m_ok, d_ok, d.code, d.doc_type, d.revision, d.state, d.description]
            extra = [sw_values.get(d.code, {}).get(p, "") for p in props]
            row_tag = self.app._state_row_tag(d.state)
            
            rows.append({"values": (base + extra), "tags": (row_tag,) if row_tag else ()})
        
        self.rc_table.set_rows(rows)
        self._on_rc_select(None)
    
    # ========== WORKFLOW PANEL (RIGHT PANEL) ==========
    
    def _load_selected_doc(self) -> "Document | None":
        """Carica documento dal codice selezionato (workflow o tabella)."""
        code = self.wf_code_var.get().strip()
        if not code:
            code = self.app._get_table_selected_code()
            if code:
                self.wf_code_var.set(code)
        if not code:
            return None
        return self.store.get_document(code)
    
    def _update_workflow_buttons(self, doc: "Document | None") -> None:
        """Abilita/disabilita pulsanti workflow in base allo stato documento."""
        def _set(btn_name: str, enabled: bool) -> None:
            btn = getattr(self, btn_name, None)
            if btn is not None:
                btn.configure(state=("normal" if enabled else "disabled"))
        
        if not doc:
            _set("wf_btn_release", False)
            _set("wf_btn_create_rev", False)
            _set("wf_btn_approve", False)
            _set("wf_btn_cancel", False)
            _set("wf_btn_obsolete", False)
            _set("wf_btn_restore_obs", False)
            return
        
        state = str(getattr(doc, "state", "") or "").strip().upper()
        prev_state = str(getattr(doc, "obs_prev_state", "") or "").strip().upper()
        
        _set("wf_btn_release", state == "WIP")
        _set("wf_btn_create_rev", state == "REL")
        _set("wf_btn_approve", state == "IN_REV")
        _set("wf_btn_cancel", state == "IN_REV")
        _set("wf_btn_obsolete", state in ("WIP", "REL", "IN_REV"))
        _set("wf_btn_restore_obs", state == "OBS" and prev_state in ("WIP", "REL", "IN_REV"))
    
    def refresh_workflow(self):
        """
        Aggiorna pannello workflow con info documento selezionato.
        
        Chiamato da:
        - Pulsante Carica
        - _send_rc_to_workflow
        - _switch_workspace
        - _refresh_all
        - Dopo transizioni workflow
        """
        doc = self._load_selected_doc()
        if not doc:
            self.wf_state_var.set("")
            self.wf_info.delete("1.0", "end")
            self._update_workflow_buttons(None)
            return
        
        self._update_workflow_buttons(doc)
        self.wf_state_var.set(f"Stato: {doc.state}  Rev: {doc.revision:02d}")
        self.wf_info.delete("1.0", "end")
        
        def _shown_path(path_s: str) -> str:
            try:
                return str(path_s or "") if (path_s and Path(path_s).is_file()) else ""
            except Exception:
                return ""
        
        # Info documento
        self.wf_info.insert("end", f"Codice: {doc.code}\n")
        self.wf_info.insert("end", f"Tipo: {doc.doc_type}\n")
        self.wf_info.insert("end", f"MMM/GGGG: {doc.mmm}/{doc.gggg}\n")
        self.wf_info.insert("end", f"Seq: {doc.seq:04d}\n")
        self.wf_info.insert("end", f"VVV: {doc.vvv}\n")
        self.wf_info.insert("end", f"Descrizione: {doc.description}\n")
        
        if getattr(doc, "obs_prev_state", ""):
            self.wf_info.insert("end", f"Stato precedente OBS: {doc.obs_prev_state}\n")
        
        self.wf_info.insert("end", "\n")
        
        # File path per stato
        self.wf_info.insert("end", f"MODEL WIP: {_shown_path(doc.file_wip_path)}\n")
        self.wf_info.insert("end", f"MODEL REL: {_shown_path(doc.file_rel_path)}\n")
        self.wf_info.insert("end", f"MODEL INREV: {_shown_path(doc.file_inrev_path)}\n\n")
        
        self.wf_info.insert("end", f"DRW WIP: {_shown_path(doc.file_wip_drw_path)}\n")
        self.wf_info.insert("end", f"DRW REL: {_shown_path(doc.file_rel_drw_path)}\n")
        self.wf_info.insert("end", f"DRW INREV: {_shown_path(doc.file_inrev_drw_path)}\n")
        
        # File REV storici
        rev_models, rev_drws = self.app._list_rev_files_for_doc(doc)
        
        self.wf_info.insert("end", "\nMODEL REV:\n")
        if rev_models:
            for p in rev_models:
                self.wf_info.insert("end", f"  {p}\n")
        else:
            self.wf_info.insert("end", "  (nessuno)\n")
        
        self.wf_info.insert("end", "DRW REV:\n")
        if rev_drws:
            for p in rev_drws:
                self.wf_info.insert("end", f"  {p}\n")
        else:
            self.wf_info.insert("end", "  (nessuno)\n")
        
        # Note cambio stato
        self.wf_info.insert("end", "\nNOTE CAMBIO STATO:\n")
        try:
            notes = self.store.list_state_notes(doc.code, limit=300)
        except Exception:
            notes = []
        
        if not notes:
            self.wf_info.insert("end", "(nessuna nota)\n")
            return
        
        for n in notes:
            ts = str(n.get("created_at", "")).replace("T", " ")
            ev = str(n.get("event_type", "")).strip().upper()
            st_from = str(n.get("from_state", "")).strip().upper()
            st_to = str(n.get("to_state", "")).strip().upper()
            rb = int(n.get("rev_before", 0))
            ra = int(n.get("rev_after", 0))
            note_text = str(n.get("note", "")).strip()
            
            self.wf_info.insert("end", f"{ts} | {ev} | {st_from} -> {st_to} | REV {rb:02d}->{ra:02d}\n")
            if note_text:
                self.wf_info.insert("end", f"{note_text}\n")
            self.wf_info.insert("end", "\n")
