"""Tab Codifica - Generazione documenti (MACHINE, GROUP, PART, ASSY)."""

from __future__ import annotations
from typing import TYPE_CHECKING
from pathlib import Path
import tkinter as tk
from tkinter import filedialog
import customtkinter as ctk

from .base_tab import BaseTab, warn, info
from pdm_sw.models import Document
from pdm_sw.codegen import build_code, build_machine_code, build_group_code
from pdm_sw.archive import archive_dirs, archive_dirs_for_machine, archive_dirs_for_group, model_path, drw_path, safe_copy, set_readonly

if TYPE_CHECKING:
    from pdm_sw.config import AppConfig
    from pdm_sw.store import Store


class TabCodifica(BaseTab):
    """Tab per la generazione di documenti: macchine, gruppi, parti e assiemi.
    
    Funzionalità:
    - Selezione tipo documento tramite radio button
    - Parametrizzazione (MMM, GGGG, variante)
    - Preview del prossimo codice disponibile
    - Creazione solo codice o con file SolidWorks (modello ± disegno)
    - Import file esistenti
    """
    
    def __init__(self, parent_frame, app, cfg: AppConfig, store: Store, session: dict):
        """Inizializza il tab Codifica.
        
        Args:
            parent_frame: Frame tkinter genitore dove costruire l'UI
            app: Riferimento all'applicazione principale
            cfg: Configurazione applicazione
            store: Store database
            session: Dati sessione utente
        """
        super().__init__(app, cfg, store, session)
        self.parent = parent_frame
        
        # Widget variables
        self.doc_type_var = None
        self.file_mode_var = None
        self.mmm_var = None
        self.gggg_var = None
        self.use_vvv_var = None
        self.vvv_choice_var = None
        self.desc_var = None
        self.link_file_var = None
        self.link_auto_drw_var = None
        self.create_checkout_var = None
        
        # Widget references
        self.mmm_menu = None
        self.gggg_menu = None
        self.vvv_check = None
        self.vvv_menu = None
        self.preview_label = None
        
        self._build_ui()
    
    def _build_ui(self):
        """Costruisce l'interfaccia del tab."""
        frame = ctk.CTkFrame(self.parent)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Variabili
        self.doc_type_var = tk.StringVar(value="PART")
        self.file_mode_var = tk.StringVar(value="code_only")  # code_only, model, model_drw
        self.mmm_var = tk.StringVar(value="")
        self.gggg_var = tk.StringVar(value="")
        self.use_vvv_var = tk.BooleanVar(value=self.cfg.code.include_vvv_by_default)
        self.vvv_choice_var = tk.StringVar(value=(self.cfg.code.vvv_presets[0] if self.cfg.code.vvv_presets else "V01"))
        self.desc_var = tk.StringVar(value="")
        self.link_file_var = tk.StringVar(value="")
        self.link_auto_drw_var = tk.BooleanVar(value=True)
        self.create_checkout_var = tk.BooleanVar(value=False)

        # --- TIPO DOCUMENTO ---
        type_frame = ctk.CTkFrame(frame)
        type_frame.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(type_frame, text="TIPO DOCUMENTO", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=8, pady=6)
        
        radio_frame = ctk.CTkFrame(type_frame, fg_color="transparent")
        radio_frame.pack(fill="x", padx=8, pady=4)
        
        ctk.CTkRadioButton(radio_frame, text="Macchina (MMM-V####) → crea ASM", variable=self.doc_type_var, value="MACHINE", command=self._on_doc_type_change).pack(anchor="w", pady=2)
        ctk.CTkRadioButton(radio_frame, text="Gruppo (MMM_GGGG-V####) → crea ASM", variable=self.doc_type_var, value="GROUP", command=self._on_doc_type_change).pack(anchor="w", pady=2)
        ctk.CTkRadioButton(radio_frame, text="Parte (MMM_GGGG-0001) → crea PRT", variable=self.doc_type_var, value="PART", command=self._on_doc_type_change).pack(anchor="w", pady=2)
        ctk.CTkRadioButton(radio_frame, text="Assieme (MMM_GGGG-9999) → crea ASM", variable=self.doc_type_var, value="ASSY", command=self._on_doc_type_change).pack(anchor="w", pady=2)

        # --- PARAMETRI ---
        params_frame = ctk.CTkFrame(frame)
        params_frame.pack(fill="x", pady=(0, 10))
        
        params_grid = ctk.CTkFrame(params_frame, fg_color="transparent")
        params_grid.pack(fill="x", padx=8, pady=8)
        
        ctk.CTkLabel(params_grid, text="MMM").grid(row=0, column=0, padx=6, pady=6, sticky="w")
        self.mmm_menu = ctk.CTkOptionMenu(params_grid, variable=self.mmm_var, values=[""], width=140, command=lambda _: self._refresh_group_menu())
        self.mmm_menu.grid(row=0, column=1, padx=6, pady=6, sticky="w")
        
        ctk.CTkLabel(params_grid, text="GGGG").grid(row=0, column=2, padx=6, pady=6, sticky="w")
        self.gggg_menu = ctk.CTkOptionMenu(params_grid, variable=self.gggg_var, values=[""], width=140, command=lambda _: self._refresh_preview())
        self.gggg_menu.grid(row=0, column=3, padx=6, pady=6, sticky="w")
        
        self.vvv_check = ctk.CTkCheckBox(params_grid, text="Variante", variable=self.use_vvv_var, command=self._refresh_preview)
        self.vvv_check.grid(row=0, column=4, padx=6, pady=6, sticky="w")
        
        self.vvv_menu = ctk.CTkOptionMenu(params_grid, variable=self.vvv_choice_var, values=self.cfg.code.vvv_presets or ["V01"], width=120, command=lambda _: self._refresh_preview())
        self.vvv_menu.grid(row=0, column=5, padx=6, pady=6, sticky="w")
        
        ctk.CTkLabel(params_grid, text="Descrizione").grid(row=1, column=0, padx=6, pady=6, sticky="w")
        ctk.CTkEntry(params_grid, textvariable=self.desc_var, width=500).grid(row=1, column=1, columnspan=5, padx=6, pady=6, sticky="ew")
        
        params_grid.grid_columnconfigure(5, weight=1)

        # File esistente (opzionale)
        link_frame = ctk.CTkFrame(params_frame, fg_color="transparent")
        link_frame.pack(fill="x", padx=8, pady=(0, 8))
        ctk.CTkLabel(link_frame, text="File esistente (opz)").pack(side="left", padx=6)
        ctk.CTkEntry(link_frame, textvariable=self.link_file_var, width=350).pack(side="left", padx=6)
        ctk.CTkButton(link_frame, text="Sfoglia", width=90, command=self._browse_link_file).pack(side="left", padx=4)
        ctk.CTkButton(link_frame, text="Pulisci", width=80, command=self._clear_link_file).pack(side="left", padx=4)
        ctk.CTkCheckBox(link_frame, text="Importa anche DRW", variable=self.link_auto_drw_var).pack(side="left", padx=12)

        # --- CREAZIONE FILE SOLIDWORKS ---
        file_frame = ctk.CTkFrame(frame)
        file_frame.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(file_frame, text="CREAZIONE FILE SOLIDWORKS", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=8, pady=6)
        
        file_radio_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        file_radio_frame.pack(fill="x", padx=8, pady=4)
        
        ctk.CTkRadioButton(file_radio_frame, text="Solo codice (no file)", variable=self.file_mode_var, value="code_only").pack(anchor="w", pady=2)
        ctk.CTkRadioButton(file_radio_frame, text="Modello (PRT/ASM automatico)", variable=self.file_mode_var, value="model").pack(anchor="w", pady=2)
        ctk.CTkRadioButton(file_radio_frame, text="Modello + Disegno", variable=self.file_mode_var, value="model_drw").pack(anchor="w", pady=2)

        # --- AZIONI ---
        actions_frame = ctk.CTkFrame(frame)
        actions_frame.pack(fill="x", pady=(10, 0))
        
        left_actions = ctk.CTkFrame(actions_frame, fg_color="transparent")
        left_actions.pack(side="left", padx=8, pady=8)
        ctk.CTkButton(left_actions, text="PROSSIMO CODICE", width=160, command=self._show_next_code).pack(side="left", padx=6)
        
        self.preview_label = ctk.CTkLabel(left_actions, text="", font=ctk.CTkFont(size=18, weight="bold"), text_color="#2E7D32")
        self.preview_label.pack(side="left", padx=12)
        
        right_actions = ctk.CTkFrame(actions_frame, fg_color="transparent")
        right_actions.pack(side="right", padx=8, pady=8)
        ctk.CTkCheckBox(
            right_actions,
            text="Crea in CHECK-OUT",
            variable=self.create_checkout_var,
        ).pack(anchor="e", pady=(0, 6))
        ctk.CTkButton(right_actions, text="GENERA", width=180, height=40, font=ctk.CTkFont(size=16, weight="bold"), fg_color="#27AE60", hover_color="#229954", command=self._generate_document).pack()

        self._on_doc_type_change()
        self._refresh_preview()
    
    # === REFRESH METHODS ===
    
    def refresh_machine_menus(self):
        """Aggiorna il menu a tendina delle macchine."""
        machines = [m for m, _ in self.store.list_machines()]
        if not machines:
            machines = [""]
        self.mmm_menu.configure(values=machines)
        if self.mmm_var.get() not in machines:
            self.mmm_var.set(machines[0])
        self._refresh_group_menu()
    
    def _refresh_group_menu(self):
        """Aggiorna il menu a tendina dei gruppi in base alla macchina selezionata."""
        mmm = self.mmm_var.get()
        groups = [g for g, _ in self.store.list_groups(mmm)] if mmm else [""]
        if not groups:
            groups = [""]
        self.gggg_menu.configure(values=groups)
        if self.gggg_var.get() not in groups:
            self.gggg_var.set(groups[0])
        self._refresh_preview()
    
    def refresh_vvv_menu(self):
        """Aggiorna il menu a tendina delle varianti."""
        vals = self.cfg.code.vvv_presets or ["V01"]
        self.vvv_menu.configure(values=vals)
        if self.vvv_choice_var.get() not in vals:
            self.vvv_choice_var.set(vals[0])
        self._refresh_preview()
    
    def _on_doc_type_change(self):
        """Abilita/disabilita controlli in base al tipo documento selezionato."""
        doc_type = self.doc_type_var.get()
        
        # MACHINE: serve solo MMM
        # GROUP: serve MMM + GGGG
        # PART/ASSY: serve MMM + GGGG + opzionale variante
        
        if doc_type == "MACHINE":
            # Disabilita GGGG e variante
            self.gggg_menu.configure(state="disabled")
            self.vvv_check.configure(state="disabled")
            self.vvv_menu.configure(state="disabled")
            self.use_vvv_var.set(False)
        elif doc_type == "GROUP":
            # Abilita GGGG, disabilita variante
            self.gggg_menu.configure(state="normal")
            self.vvv_check.configure(state="disabled")
            self.vvv_menu.configure(state="disabled")
            self.use_vvv_var.set(False)
        else:  # PART/ASSY
            # Abilita tutto
            self.gggg_menu.configure(state="normal")
            self.vvv_check.configure(state="normal")
            self.vvv_menu.configure(state="normal")
        
        self._refresh_preview()
    
    def _refresh_preview(self):
        """Mostra preview del codice."""
        doc_type = self.doc_type_var.get()
        mmm = self.mmm_var.get() or "MMM"
        gggg = self.gggg_var.get() or "GGGG"
        
        if doc_type == "MACHINE":
            code = build_machine_code(self.cfg, mmm, 1)
            code = code.replace("V0001", "V0000")  # Preview placeholder
        elif doc_type == "GROUP":
            code = build_group_code(self.cfg, mmm, gggg, 1)
            code = code.replace("V0001", "V0000")
        else:  # PART/ASSY
            seq = 1
            vvv = self.vvv_choice_var.get() if self.use_vvv_var.get() else ""
            code = build_code(self.cfg, mmm, gggg, seq, vvv=vvv, force_vvv=self.use_vvv_var.get())
            code = code.replace("0001", "0000")
        
        self.preview_label.configure(text=f"Preview: {code}")
    
    def _show_next_code(self):
        """Mostra il prossimo codice disponibile senza consumare la sequenza."""
        doc_type = self.doc_type_var.get()
        mmm = (self.mmm_var.get() or "").strip()
        
        if not mmm:
            warn("Seleziona MMM.")
            return
        
        try:
            if doc_type == "MACHINE":
                # Peek ver_counters per MACHINE
                row = self.store.conn.execute(
                    "SELECT next_ver FROM ver_counters WHERE mmm=? AND gggg='' AND doc_type='MACHINE';",
                    (mmm,)
                ).fetchone()
                seq = int(row["next_ver"]) if row else 1
                code = build_machine_code(self.cfg, mmm, seq)
            
            elif doc_type == "GROUP":
                gggg = (self.gggg_var.get() or "").strip()
                if not gggg:
                    warn("Seleziona GGGG.")
                    return
                row = self.store.conn.execute(
                    "SELECT next_ver FROM ver_counters WHERE mmm=? AND gggg=? AND doc_type='GROUP';",
                    (mmm, gggg)
                ).fetchone()
                seq = int(row["next_ver"]) if row else 1
                code = build_group_code(self.cfg, mmm, gggg, seq)
            
            else:  # PART/ASSY
                gggg = (self.gggg_var.get() or "").strip()
                if not gggg:
                    warn("Seleziona GGGG.")
                    return
                vvv = (self.vvv_choice_var.get() or "").strip().upper() if self.use_vvv_var.get() else ""
                seq = self.store.peek_seq(mmm, gggg, vvv, doc_type)
                code = build_code(self.cfg, mmm, gggg, seq, vvv=vvv, force_vvv=self.use_vvv_var.get())
            
            self.preview_label.configure(text=f"Prossimo: {code}")
        except Exception as e:
            warn(f"Errore calcolo prossimo codice: {e}")
    
    # === FILE LINK METHODS ===
    
    def _browse_link_file(self):
        """Apre dialog per selezionare un file SolidWorks esistente."""
        path = filedialog.askopenfilename(
            title='Seleziona file SolidWorks',
            filetypes=[('SolidWorks', '*.sldprt *.sldasm *.slddrw'), ('Tutti i file', '*.*')]
        )
        if path:
            self.link_file_var.set(path)
    
    def _clear_link_file(self):
        """Pulisce il campo file esistente."""
        self.link_file_var.set('')
    
    def _import_linked_files_to_wip(self, doc: Document, src_path: str) -> None:
        """Importa un file esistente e il relativo DRW nella cartella WIP dell'archivio.
        
        Args:
            doc: Documento PDM di destinazione
            src_path: Path del file sorgente da importare
        """
        if not self.cfg.solidworks.archive_root:
            raise ValueError('Archivio non configurato (SolidWorks > Archivio).')
        p = Path(src_path)
        if not p.exists():
            raise ValueError('File selezionato non trovato.')

        ext = p.suffix.lower()
        # determina sorgente modello
        src_model = p
        if ext == '.slddrw':
            # prova a trovare modello con stesso nome
            cand_prt = p.with_suffix('.sldprt')
            cand_asm = p.with_suffix('.sldasm')
            if doc.doc_type == 'PART' and cand_prt.exists():
                src_model = cand_prt
            elif doc.doc_type == 'ASSY' and cand_asm.exists():
                src_model = cand_asm
            elif cand_prt.exists():
                src_model = cand_prt
            elif cand_asm.exists():
                src_model = cand_asm
            else:
                raise ValueError('Hai selezionato un DRW ma non trovo il modello (.sldprt/.sldasm) con stesso nome nella stessa cartella.')
        # check coerenza tipo
        if src_model.suffix.lower() == '.sldprt' and doc.doc_type != 'PART':
            raise ValueError("Il file selezionato è una PARTE (.sldprt) ma il tipo scelto è ASSY.")
        if src_model.suffix.lower() == '.sldasm' and doc.doc_type != 'ASSY':
            raise ValueError("Il file selezionato è un ASSIEME (.sldasm) ma il tipo scelto è PART.")

        wip, rel, inrev, rev = archive_dirs(self.cfg.solidworks.archive_root, doc.mmm, doc.gggg)
        dst_model = model_path(wip, doc.code, doc.doc_type)
        if dst_model.exists():
            raise ValueError("Esiste già un file modello in archivio con questo codice (WIP).")
        safe_copy(src_model, dst_model)
        set_readonly(dst_model, readonly=False)
        self.store.update_document(doc.code, file_wip_path=str(dst_model))

        # DRW: stesso nome del modello, stessa cartella
        if bool(self.link_auto_drw_var.get()):
            src_drw = src_model.with_suffix('.slddrw')
            if src_drw.exists():
                dst_drw = drw_path(wip, doc.code)
                if not dst_drw.exists():
                    safe_copy(src_drw, dst_drw)
                    set_readonly(dst_drw, readonly=False)
                self.store.update_document(doc.code, file_wip_drw_path=str(dst_drw))
    
    # === DOCUMENT GENERATION ===
    
    def _generate_document(self):
        """Funzione unificata per generare qualsiasi tipo di documento."""
        doc_type = self.doc_type_var.get()
        file_mode = self.file_mode_var.get()
        mmm = self.mmm_var.get().strip()
        
        # Validazione parametri base
        if not mmm:
            warn("Seleziona MMM.")
            return
        
        desc = self._require_desc_upper(self.desc_var.get(), what="descrizione")
        if desc is None:
            return
        self.desc_var.set(desc)
        
        # Validazione specifica per tipo
        gggg = ""
        vvv = ""
        code = ""
        seq = 0
        action_log = ""
        
        try:
            if doc_type == "MACHINE":
                # MACHINE: solo MMM necessario
                seq = self.store.allocate_ver_seq(mmm, "", "MACHINE")
                code = build_machine_code(self.cfg, mmm, seq)
                
                # Paths
                wip_path = ""
                wip_drw_path = ""
                if self.cfg.solidworks.archive_root:
                    wip, rel, inrev, rev = archive_dirs_for_machine(self.cfg.solidworks.archive_root, mmm)
                    wip_path = str(model_path(wip, code, "MACHINE"))
                    wip_drw_path = str(drw_path(wip, code)) if file_mode == "model_drw" else ""
                
                doc = Document(
                    id=0, code=code, doc_type="MACHINE", mmm=mmm, gggg="", seq=seq, vvv="",
                    revision=0, state="WIP", description=desc,
                    file_wip_path=wip_path, file_rel_path="", file_inrev_path="",
                    file_wip_drw_path=wip_drw_path, file_rel_drw_path="", file_inrev_drw_path="",
                    created_at="", updated_at=""
                )
                action_log = "CREATE_MACHINE"
                
            elif doc_type == "GROUP":
                # GROUP: MMM + GGGG necessari
                gggg = self.gggg_var.get().strip()
                if not gggg:
                    warn("Seleziona GGGG.")
                    return
                
                seq = self.store.allocate_ver_seq(mmm, gggg, "GROUP")
                code = build_group_code(self.cfg, mmm, gggg, seq)
                
                # Paths
                wip_path = ""
                wip_drw_path = ""
                if self.cfg.solidworks.archive_root:
                    wip, rel, inrev, rev = archive_dirs_for_group(self.cfg.solidworks.archive_root, mmm, gggg)
                    wip_path = str(model_path(wip, code, "GROUP"))
                    wip_drw_path = str(drw_path(wip, code)) if file_mode == "model_drw" else ""
                
                doc = Document(
                    id=0, code=code, doc_type="GROUP", mmm=mmm, gggg=gggg, seq=seq, vvv="",
                    revision=0, state="WIP", description=desc,
                    file_wip_path=wip_path, file_rel_path="", file_inrev_path="",
                    file_wip_drw_path=wip_drw_path, file_rel_drw_path="", file_inrev_drw_path="",
                    created_at="", updated_at=""
                )
                action_log = "CREATE_GROUP"
                
            else:  # PART or ASSY
                # PART/ASSY: MMM + GGGG + opzionale variante
                gggg = self.gggg_var.get().strip()
                if not gggg:
                    warn("Seleziona GGGG.")
                    return
                
                vvv = self.vvv_choice_var.get().strip() if self.use_vvv_var.get() else ""
                vvv_key = self.cfg.code.segments["VVV"].normalize_value(vvv) if vvv else ""
                seq = self.store.allocate_seq(mmm, gggg, vvv_key, doc_type)
                vvv = vvv_key
                code = build_code(self.cfg, mmm, gggg, seq, vvv=vvv, force_vvv=self.use_vvv_var.get())
                
                # Paths
                wip_path = ""
                wip_drw_path = ""
                if self.cfg.solidworks.archive_root:
                    wip, rel, inrev, rev = archive_dirs(self.cfg.solidworks.archive_root, mmm, gggg)
                    wip_path = str(model_path(wip, code, doc_type))
                    wip_drw_path = str(drw_path(wip, code)) if file_mode == "model_drw" else ""
                
                doc = Document(
                    id=0, code=code, doc_type=doc_type, mmm=mmm, gggg=gggg, seq=seq, vvv=vvv,
                    revision=0, state="WIP", description=desc,
                    file_wip_path=wip_path, file_rel_path="", file_inrev_path="",
                    file_wip_drw_path=wip_drw_path, file_rel_drw_path="", file_inrev_drw_path="",
                    created_at="", updated_at=""
                )
                action_log = "CREATE_CODE"
            
            # Crea documento in DB
            self.store.add_document(doc)

            if bool(self.create_checkout_var.get()):
                if not self.app._checkout_document(code, show_feedback=False, refresh_ui=False):
                    warn("Documento creato, ma CHECK-OUT automatico non riuscito.")
                    self.app._log_activity(action_log, code=code, status="WARN", message="Creato ma checkout automatico fallito")
                    self.app._refresh_all()
                    return
            
            # Import file esistente se specificato
            link_file = self.link_file_var.get().strip()
            if link_file and doc_type in ("PART", "ASSY"):
                try:
                    self._import_linked_files_to_wip(doc, link_file)
                except Exception as imp_e:
                    warn(f"Codice creato, ma import file fallito: {imp_e}")
                    self.app._log_activity(action_log, code=code, status="WARN", message=f"Creato ma import fallito: {imp_e}")
            
            # Crea file SolidWorks se richiesto
            if file_mode in ("model", "model_drw"):
                if link_file and doc_type in ("PART", "ASSY"):
                    warn(f"Hai selezionato un file esistente: verrà importato, non creato da template.")
                else:
                    # Delega alla funzione _create_files_for_code dell'app principale
                    self.app._create_files_for_code(
                        code,
                        create_drw=(file_mode == "model_drw"),
                        only_missing=False,
                        require_checkout=False,
                    )
            
            info(f"Documento creato: {code}")
            self.app._log_activity(action_log, code=code, status="OK", message=f"Creato {doc_type}")
            self.app._refresh_all()
            
        except Exception as e:
            self.app._log_activity(action_log if action_log else "CREATE_ERROR", 
                             code=code if code else "", status="ERROR", message=str(e))
            warn(f"Errore creazione documento: {e}")
