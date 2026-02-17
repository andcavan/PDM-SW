# -*- coding: utf-8 -*-
"""
Tab SolidWorks
Configurazione archivio, template, property mapping, test connessione
"""
import tkinter as tk
from tkinter import filedialog
import customtkinter as ctk
from pdm_sw.sw_integration import test_solidworks_connection
from pdm_sw.macro_publish import publish_macro
from .base_tab import BaseTab, warn, info


class TabSolidWorks(BaseTab):
    """Tab configurazione SolidWorks e macro bootstrap."""

    def __init__(self, parent_frame, app, cfg, store, session):
        super().__init__(app, cfg, store, session)
        self.root = parent_frame
        
        # Widget references
        self.archive_root_var = None
        self.tpl_part_var = None
        self.tpl_assy_var = None
        self.tpl_drw_var = None
        self.sw_desc_prop_var = None
        self.sw_map_rows = []
        self.sw_read_rows = []
        self.sw_status = None
        
        self._build_ui()

    def _build_ui(self):
        """Costruisce l'interfaccia del tab."""
        frame = ctk.CTkScrollableFrame(self.root)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(
            frame,
            text="Impostazioni SolidWorks",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", pady=(0, 10))

        def row(label: str, var: tk.StringVar, browse: bool = True, is_file: bool = False):
            r = ctk.CTkFrame(frame)
            r.pack(fill="x", pady=6)
            ctk.CTkLabel(r, text=label, width=170, anchor="w").pack(side="left", padx=(8, 6))
            ctk.CTkEntry(r, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
            if browse:
                def _pick():
                    if is_file:
                        p = filedialog.askopenfilename()
                    else:
                        p = filedialog.askdirectory()
                    if p:
                        var.set(p)
                ctk.CTkButton(r, text="Sfoglia", width=90, command=_pick).pack(side="left", padx=6)

        self.archive_root_var = tk.StringVar(value=self.cfg.solidworks.archive_root)
        self.tpl_part_var = tk.StringVar(value=self.cfg.solidworks.template_part)
        self.tpl_assy_var = tk.StringVar(value=self.cfg.solidworks.template_assembly)
        self.tpl_drw_var = tk.StringVar(value=self.cfg.solidworks.template_drawing)

        row("Archivio (root)", self.archive_root_var, browse=True, is_file=False)
        row("Template PART", self.tpl_part_var, browse=True, is_file=True)
        row("Template ASSY", self.tpl_assy_var, browse=True, is_file=True)
        row("Template DRW", self.tpl_drw_var, browse=True, is_file=True)

        # ---- Mappatura proprietà: PDM -> SolidWorks
        ctk.CTkLabel(
            frame,
            text="Mappatura proprietà (PDM -> SolidWorks)",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", pady=(18, 6))
        
        ctk.CTkLabel(
            frame,
            text="Definisci quali proprietà personalizzate scrivere nei file SolidWorks. "
                 "Ogni riga collega un campo PDM a una proprietà custom SolidWorks. "
                 "Puoi aggiungere e cancellare righe liberamente.",
            text_color="#777777",
            wraplength=820,
            justify="left"
        ).pack(anchor="w", pady=(0, 8))

        hdr = ctk.CTkFrame(frame, fg_color="transparent")
        hdr.pack(fill="x", pady=(0, 2))
        ctk.CTkLabel(hdr, text="Campo PDM", width=180, anchor="w").pack(side="left", padx=(8, 6))
        ctk.CTkLabel(hdr, text="Proprietà SolidWorks (nome)", anchor="w").pack(side="left", padx=6)

        map_list = ctk.CTkScrollableFrame(frame, height=220)
        map_list.pack(fill="x", pady=(0, 8))

        self.sw_map_rows = []
        pdm_fields = self._pdm_fields_for_mapping()

        def _remove_sw_map_row(row_frame):
            self.sw_map_rows = [r for r in self.sw_map_rows if r.get("frame") is not row_frame]
            try:
                row_frame.destroy()
            except Exception:
                pass

        def _add_sw_map_row(pdm_field: str = "code", sw_prop: str = ""):
            rf = ctk.CTkFrame(map_list, fg_color="transparent")
            rf.pack(fill="x", pady=3)

            field_var = tk.StringVar(value=(pdm_field if pdm_field in pdm_fields else "code"))
            sw_var = tk.StringVar(value=(sw_prop or ""))

            opt = ctk.CTkOptionMenu(rf, variable=field_var, values=pdm_fields, width=180)
            opt.pack(side="left", padx=(8, 6))
            ctk.CTkEntry(rf, textvariable=sw_var).pack(side="left", fill="x", expand=True, padx=6)
            ctk.CTkButton(rf, text="X", width=34, command=lambda: _remove_sw_map_row(rf)).pack(side="left", padx=6)

            self.sw_map_rows.append({"frame": rf, "field_var": field_var, "sw_var": sw_var, "opt": opt})

        def _load_sw_map_rows():
            mappings = []
            try:
                mappings = list(getattr(self.cfg.solidworks, "property_mappings", []) or [])
            except Exception:
                mappings = []
            
            if not mappings:
                legacy = dict(getattr(self.cfg.solidworks, "property_map", {}) or {})
                if legacy:
                    mappings = [{"pdm_field": k, "sw_prop": v} for k, v in legacy.items()]
                else:
                    d = self._default_sw_property_map()
                    mappings = [{"pdm_field": k, "sw_prop": v} for k, v in d.items()]

            for r in list(self.sw_map_rows):
                try:
                    r["frame"].destroy()
                except Exception:
                    pass
            self.sw_map_rows.clear()

            for it in mappings:
                pf = str(it.get("pdm_field", it.get("field", "code")) or "code")
                if pf == "description":
                    # descrizione ora gestita da SolidWorks
                    try:
                        self.cfg.solidworks.description_prop = str(it.get("sw_prop", it.get("sw", "DESCRIZIONE")) or "DESCRIZIONE")
                        self.sw_desc_prop_var.set(self.cfg.solidworks.description_prop)
                    except Exception:
                        pass
                    continue
                sp = str(it.get("sw_prop", it.get("sw", "")) or "")
                _add_sw_map_row(pf, sp)

        _load_sw_map_rows()

        map_btns = ctk.CTkFrame(frame, fg_color="transparent")
        map_btns.pack(fill="x", pady=(0, 10))

        def _reset_map_defaults():
            d = self._default_sw_property_map()
            for r in list(self.sw_map_rows):
                try:
                    r["frame"].destroy()
                except Exception:
                    pass
            self.sw_map_rows.clear()
            for k, v in d.items():
                _add_sw_map_row(k, v)

        ctk.CTkButton(
            map_btns,
            text="Aggiungi proprietà",
            width=160,
            command=lambda: _add_sw_map_row()
        ).pack(side="left", padx=8)
        
        ctk.CTkButton(
            map_btns,
            text="Ripristina default",
            width=160,
            command=_reset_map_defaults
        ).pack(side="left", padx=8)

        # ---- Descrizione (gestita da SolidWorks)
        ctk.CTkLabel(
            frame,
            text="Descrizione (gestita da SolidWorks)",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", pady=(10, 6))
        
        ctk.CTkLabel(
            frame,
            text="La descrizione viene inserita in PDM alla creazione del codice (seed iniziale). "
                 "Alla creazione file viene scritta nel file SolidWorks e da quel momento è gestita da SolidWorks. "
                 "Il PDM la legge e la visualizza nelle tabelle.",
            text_color="#777777",
            wraplength=820,
            justify="left"
        ).pack(anchor="w", pady=(0, 6))

        self.sw_desc_prop_var = tk.StringVar(
            value=(getattr(self.cfg.solidworks, "description_prop", "DESCRIZIONE") or "DESCRIZIONE")
        )

        desc_row = ctk.CTkFrame(frame, fg_color="transparent")
        desc_row.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(desc_row, text="Nome proprietà SW", width=180, anchor="w").pack(side="left", padx=(8, 6))
        ctk.CTkEntry(desc_row, textvariable=self.sw_desc_prop_var).pack(side="left", fill="x", expand=True, padx=6)

        # ---- Proprietà custom da leggere (SW -> PDM)
        ctk.CTkLabel(
            frame,
            text="Proprietà custom da leggere (SW -> PDM)",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", pady=(10, 6))
        
        ctk.CTkLabel(
            frame,
            text="Elenca le proprietà custom SolidWorks che il PDM deve leggere (oltre la descrizione). "
                 "Questi valori vengono aggiornati dopo i cambi di stato e tramite 'Forza SW->PDM' in Consultazione.",
            text_color="#777777",
            wraplength=820,
            justify="left"
        ).pack(anchor="w", pady=(0, 8))

        read_hdr = ctk.CTkFrame(frame, fg_color="transparent")
        read_hdr.pack(fill="x", pady=(0, 2))
        ctk.CTkLabel(read_hdr, text="Proprietà SolidWorks (nome)", anchor="w").pack(side="left", padx=(8, 6))

        read_list = ctk.CTkScrollableFrame(frame, height=160)
        read_list.pack(fill="x", pady=(0, 8))

        self.sw_read_rows = []

        def _remove_sw_read_row(rf):
            self.sw_read_rows = [r for r in self.sw_read_rows if r.get("frame") is not rf]
            try:
                rf.destroy()
            except Exception:
                pass

        def _add_sw_read_row(prop_name: str = ""):
            rf = ctk.CTkFrame(read_list, fg_color="transparent")
            rf.pack(fill="x", pady=3)
            v = tk.StringVar(value=(prop_name or ""))
            ctk.CTkEntry(rf, textvariable=v).pack(side="left", fill="x", expand=True, padx=(8, 6))
            ctk.CTkButton(rf, text="X", width=34, command=lambda: _remove_sw_read_row(rf)).pack(side="left", padx=6)
            self.sw_read_rows.append({"frame": rf, "var": v})

        try:
            rp = list(getattr(self.cfg.solidworks, "read_properties", []) or [])
        except Exception:
            rp = []
        for p in rp:
            _add_sw_read_row(str(p))

        read_btns = ctk.CTkFrame(frame, fg_color="transparent")
        read_btns.pack(fill="x", pady=(0, 10))
        ctk.CTkButton(
            read_btns,
            text="Aggiungi proprietà",
            width=180,
            command=lambda: _add_sw_read_row("")
        ).pack(side="left", padx=8)

        # Pulsanti azione
        btns = ctk.CTkFrame(frame, fg_color="transparent")
        btns.pack(fill="x", pady=10)
        
        ctk.CTkButton(
            btns,
            text="Salva impostazioni",
            command=self._save_sw_config
        ).pack(side="left", padx=8)
        
        ctk.CTkButton(
            btns,
            text="PUBBLICA MACRO SOLIDWORKS",
            command=self._publish_sw_macro
        ).pack(side="left", padx=8)
        
        ctk.CTkButton(
            btns,
            text="Test connessione",
            command=self._test_sw
        ).pack(side="left", padx=8)
        
        self.sw_status = ctk.CTkLabel(btns, text="")
        self.sw_status.pack(side="left", padx=12)

    def _pdm_fields_for_mapping(self) -> list[str]:
        """Campi PDM core disponibili per il mapping."""
        return ["code", "revision", "state", "doc_type", "mmm", "gggg", "vvv"]

    def _default_sw_property_map(self) -> dict:
        """Mapping di default (italiano) per proprietà PDM -> SW."""
        return {
            "code": "CODICE",
            "revision": "REVISIONE",
            "state": "STATO",
            "doc_type": "TIPO_DOC",
            "mmm": "MACCHINA",
            "gggg": "GRUPPO",
            "vvv": "VARIANTE",
        }

    def _save_sw_config(self):
        """Salva configurazione SolidWorks."""
        self.cfg.solidworks.archive_root = self.archive_root_var.get().strip()
        self.cfg.solidworks.template_part = self.tpl_part_var.get().strip()
        self.cfg.solidworks.template_assembly = self.tpl_assy_var.get().strip()
        self.cfg.solidworks.template_drawing = self.tpl_drw_var.get().strip()
        
        # Descrizione (SW-managed)
        try:
            self.cfg.solidworks.description_prop = (
                self.sw_desc_prop_var.get() if hasattr(self, 'sw_desc_prop_var')
                else getattr(self.cfg.solidworks, 'description_prop', 'DESCRIZIONE')
            ).strip() or 'DESCRIZIONE'
        except Exception:
            self.cfg.solidworks.description_prop = 'DESCRIZIONE'
        
        # Proprietà custom da leggere (oltre descrizione)
        if hasattr(self, 'sw_read_rows'):
            rp = []
            for r in (self.sw_read_rows or []):
                try:
                    v = (r.get('var').get() if r.get('var') else '').strip()
                except Exception:
                    v = ''
                if not v:
                    continue
                rp.append(v)
            seen = set()
            clean = []
            for v in rp:
                k = str(v).strip().upper()
                if not k or k in seen:
                    continue
                seen.add(k)
                clean.append(k)
            self.cfg.solidworks.read_properties = clean
        
        # Salva mappatura proprietà (lista righe)
        if hasattr(self, "sw_map_rows"):
            mappings = []
            legacy_mp = {}
            for r in (self.sw_map_rows or []):
                try:
                    pf = (r.get("field_var").get() if r.get("field_var") else "").strip()
                    sp = (r.get("sw_var").get() if r.get("sw_var") else "").strip()
                except Exception:
                    continue
                if not sp:
                    continue
                if not pf:
                    pf = "code"
                if pf == "description":
                    continue
                mappings.append({"pdm_field": pf, "sw_prop": sp})
                if pf not in legacy_mp:
                    legacy_mp[pf] = sp
            self.cfg.solidworks.property_mappings = mappings
            self.cfg.solidworks.property_map = legacy_mp

        self.app.cfg_mgr.cfg = self.cfg
        self.app.cfg_mgr.save()
        info("Impostazioni SolidWorks salvate.")

    def _test_sw(self):
        """Testa connessione a SolidWorks."""
        st = test_solidworks_connection()
        if st.ok:
            self.sw_status.configure(text=f"OK {st.version}")
            info(f"Connessione OK. Versione: {st.version}")
        else:
            self.sw_status.configure(text="FAIL")
            warn(st.message + ("\n\n" + st.details if st.details else ""))

    def _publish_sw_macro(self) -> None:
        """Pubblica (genera) la macro di bootstrap SolidWorks + payload per la workspace corrente."""
        try:
            from pathlib import Path
            app_dir = Path(__file__).resolve().parent.parent.parent
            bas_path, payload_dir = publish_macro(app_dir, self.app.ws_id)
            info(
                "Macro SolidWorks pubblicata.\n\n"
                f"Bootstrap (.bas): {bas_path}\n"
                f"Payload: {payload_dir}\n\n"
                "Apri il file di istruzioni nella cartella SW_MACROS (INSTALL_MACRO_<workspace>.txt)."
            )
        except Exception as e:
            warn(f"Errore pubblicazione macro: {e}")
