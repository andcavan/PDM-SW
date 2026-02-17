# -*- coding: utf-8 -*-
"""
Tab Gestione Codifica
Configurazione formato codice: MMM, GGGG, VVV, progressivo
"""
import tkinter as tk
import customtkinter as ctk
from pdm_sw.config import SegmentRule
from .base_tab import BaseTab, info


class TabGestioneCodifica(BaseTab):
    """Tab di configurazione codifica e regole segmenti."""

    def __init__(self, parent_frame, app, cfg, store, session):
        super().__init__(app, cfg, store, session)
        self.root = parent_frame
        self.sep1_var = None
        self.sep2_var = None
        self.sep3_var = None
        self.include_vvv_var = None
        self.vvv_var = None
        self.seg_rule_vars = {}
        self.code_cfg_preview = None
        self.code_preview = None
        self._build_ui()

    def _build_ui(self):
        frame = ctk.CTkScrollableFrame(self.root)
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        ctk.CTkLabel(
            frame,
            text="Configurazione codifica [MMM]_[GGGG]-[VVV]-[0000]",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", pady=(0, 10))

        # Separatori
        row1 = ctk.CTkFrame(frame)
        row1.pack(fill="x", pady=6)

        self.sep1_var = tk.StringVar(value=self.cfg.code.sep1)
        self.sep2_var = tk.StringVar(value=self.cfg.code.sep2)
        self.sep3_var = tk.StringVar(value=self.cfg.code.sep3)
        self.include_vvv_var = tk.BooleanVar(value=self.cfg.code.include_vvv_by_default)

        ctk.CTkLabel(row1, text="Separatore MMM/GGGG").pack(side="left", padx=(8, 6))
        ctk.CTkEntry(row1, width=60, textvariable=self.sep1_var).pack(side="left", padx=6)
        ctk.CTkLabel(row1, text="Separatore GGGG/0000").pack(side="left", padx=(18, 6))
        ctk.CTkEntry(row1, width=60, textvariable=self.sep2_var).pack(side="left", padx=6)
        ctk.CTkLabel(row1, text="Separatore 0000/VVV").pack(side="left", padx=(18, 6))
        ctk.CTkEntry(row1, width=60, textvariable=self.sep3_var).pack(side="left", padx=6)

        # Checkbox VVV
        row2 = ctk.CTkFrame(frame)
        row2.pack(fill="x", pady=6)
        ctk.CTkCheckBox(row2, text="Includi VVV di default", variable=self.include_vvv_var).pack(side="left", padx=8)

        # Preset VVV
        row3 = ctk.CTkFrame(frame)
        row3.pack(fill="x", pady=6)
        ctk.CTkLabel(row3, text="Preset VVV (separati da virgola)").pack(side="left", padx=(8, 6))
        self.vvv_var = tk.StringVar(value=",".join(self.cfg.code.vvv_presets))
        ctk.CTkEntry(row3, textvariable=self.vvv_var).pack(side="left", fill="x", expand=True, padx=(6, 8))

        # Regole segmenti
        ctk.CTkLabel(
            frame,
            text="Regole segmenti",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", pady=(16, 6))

        seg_frame = ctk.CTkFrame(frame)
        seg_frame.pack(fill="x", pady=6)

        # Headers
        hdr = ctk.CTkFrame(seg_frame, fg_color="transparent")
        hdr.pack(fill="x", padx=8, pady=(8, 2))
        ctk.CTkLabel(hdr, text="Segmento", width=90).pack(side="left")
        ctk.CTkLabel(hdr, text="Attivo", width=70).pack(side="left")
        ctk.CTkLabel(hdr, text="Lunghezza", width=90).pack(side="left")
        ctk.CTkLabel(hdr, text="Charset", width=120).pack(side="left")
        ctk.CTkLabel(hdr, text="Case", width=120).pack(side="left")

        self.seg_rule_vars = {}
        for token in ["MMM", "GGGG", "0000", "VVV"]:
            rule = self.cfg.code.segments.get(token)
            row = ctk.CTkFrame(seg_frame, fg_color="transparent")
            row.pack(fill="x", padx=8, pady=2)

            enabled_var = tk.BooleanVar(value=(rule.enabled if rule else True))
            length_var = tk.StringVar(value=str(rule.length if rule else (4 if token == "0000" else len(token))))
            charset_var = tk.StringVar(value=(rule.charset if rule else ("NUM" if token == "0000" else "ALPHA")))
            case_var = tk.StringVar(value=(rule.case if rule else "UPPER"))

            self.seg_rule_vars[token] = {
                "enabled": enabled_var,
                "length": length_var,
                "charset": charset_var,
                "case": case_var,
            }

            ctk.CTkLabel(row, text=token, width=90).pack(side="left")
            ctk.CTkCheckBox(
                row,
                text="",
                variable=enabled_var,
                width=70,
                command=self._refresh_code_config_preview
            ).pack(side="left")
            ctk.CTkEntry(row, width=90, textvariable=length_var).pack(side="left", padx=(0, 10))
            ctk.CTkOptionMenu(
                row,
                width=120,
                values=["NUM", "ALPHA", "ALNUM"],
                variable=charset_var,
                command=lambda _=None: self._refresh_code_config_preview()
            ).pack(side="left", padx=(0, 10))
            ctk.CTkOptionMenu(
                row,
                width=120,
                values=["UPPER", "LOWER"],
                variable=case_var,
                command=lambda _=None: self._refresh_code_config_preview()
            ).pack(side="left")

        # Anteprima
        self.code_cfg_preview = ctk.CTkLabel(frame, text="")
        self.code_cfg_preview.pack(anchor="w", pady=(10, 0))
        self._refresh_code_config_preview()

        # Bottoni
        btns = ctk.CTkFrame(frame, fg_color="transparent")
        btns.pack(fill="x", pady=12)
        ctk.CTkButton(btns, text="Salva configurazione", command=self._save_code_config).pack(side="left", padx=8)
        self.code_preview = ctk.CTkLabel(btns, text="")
        self.code_preview.pack(side="left", padx=14)

    def _save_code_config(self):
        """Salva configurazione codifica nel cfg e su disco."""
        self.cfg.code.sep1 = self.sep1_var.get()
        self.cfg.code.sep2 = self.sep2_var.get()
        self.cfg.code.sep3 = self.sep3_var.get()
        self.cfg.code.include_vvv_by_default = bool(self.include_vvv_var.get())
        self.cfg.code.vvv_presets = [x.strip() for x in self.vvv_var.get().split(",") if x.strip()]

        # Segment rules
        if hasattr(self, "seg_rule_vars") and self.seg_rule_vars:
            for token, vars_ in self.seg_rule_vars.items():
                try:
                    length = int(str(vars_["length"].get()).strip() or "0")
                except Exception:
                    length = 0
                length = max(1, min(12, length)) if token != "0000" else max(1, min(6, length))
                rule = self.cfg.code.segments.get(token) or SegmentRule()
                rule.enabled = bool(vars_["enabled"].get())
                rule.length = length
                rule.charset = str(vars_["charset"].get())
                rule.case = str(vars_["case"].get())
                self.cfg.code.segments[token] = rule

        self._refresh_code_config_preview()
        self.app.cfg_mgr.cfg = self.cfg
        self.app.cfg_mgr.save()

        # Aggiorna menu VVV nella tab Codifica se presente
        if hasattr(self.app, 'tab_codifica') and hasattr(self.app.tab_codifica, 'refresh_vvv_menu'):
            self.app.tab_codifica.refresh_vvv_menu()

        info("Configurazione codifica salvata.")

    def _refresh_code_config_preview(self):
        """Anteprima formato codice in base ai parametri correnti."""
        try:
            sep1 = self.sep1_var.get()
            sep2 = self.sep2_var.get()
            sep3 = self.sep3_var.get()
            include_vvv = bool(self.include_vvv_var.get())
            vvv_sample = (self.cfg.code.vvv_presets[0] if self.cfg.code.vvv_presets else "V01")

            # Build temporary segment rules
            segs = {}
            if hasattr(self, "seg_rule_vars"):
                for token, vars_ in self.seg_rule_vars.items():
                    try:
                        length = int(str(vars_["length"].get()).strip() or "0")
                    except Exception:
                        length = 0
                    length = max(1, min(12, length)) if token != "0000" else max(1, min(6, length))
                    segs[token] = SegmentRule(
                        enabled=bool(vars_["enabled"].get()),
                        length=length,
                        charset=str(vars_["charset"].get()),
                        case=str(vars_["case"].get()),
                    )
            else:
                segs = self.cfg.code.segments

            mmm_v = segs["MMM"].normalize_value("MAC") if segs.get("MMM") and segs["MMM"].enabled else ""
            gggg_v = segs["GGGG"].normalize_value("GRUP") if segs.get("GGGG") and segs["GGGG"].enabled else ""
            seq_len = segs["0000"].length if segs.get("0000") else 4
            seq_v = str(1).zfill(seq_len) if (segs.get("0000") and segs["0000"].enabled) else ""
            code = mmm_v + sep1 + gggg_v + sep2
            if include_vvv:
                vvv_v = segs["VVV"].normalize_value(vvv_sample) if segs.get("VVV") and segs["VVV"].enabled else ""
                if vvv_v:
                    code += vvv_v + sep3 + seq_v
                else:
                    code += seq_v
            else:
                code += seq_v

            if hasattr(self, "code_cfg_preview"):
                self.code_cfg_preview.configure(text=f"Anteprima: {code}")
            if hasattr(self, "code_preview"):
                self.code_preview.configure(text=f"Anteprima: {code}")
        except Exception:
            pass
