from __future__ import annotations

def _log_line(path: str, msg: str) -> None:
    if not path:
        return
    try:
        from pathlib import Path
        p = Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        with p.open("a", encoding="utf-8") as f:
            f.write(msg.rstrip() + "\n")
    except Exception:
        pass

import argparse
import json
import os
import re as _re
from pathlib import Path
from typing import Any, Dict, Optional

import tkinter as tk
from tkinter import messagebox

try:
    import customtkinter as ctk
    _CTK_IMPORT_ERROR = None
except Exception as e:
    ctk = None
    _CTK_IMPORT_ERROR = e
from .workspace import WorkspaceManager
from .config import ConfigManager, AppConfig
from .store import Store
from .models import Document, DocType, State
from .session_context import resolve_session_context
from .codegen import build_code
from .archive import (
    archive_dirs,
    model_path,
    drw_path,
    inrev_tag,
    safe_copy,
    set_readonly,
    release_wip,
    create_inrev,
    approve_inrev,
    cancel_inrev,
    set_obsolete,
    restore_obsolete,
)
from .sw_api import (
    get_solidworks_app,
    get_custom_properties,
    set_custom_properties,
    save_existing_doc,
    open_doc,
    close_doc,
    create_model_file,
    create_drawing_file,
    save_as_doc,
)


def _warn(msg: str) -> None:
    messagebox.showwarning("PDM (Macro SolidWorks)", msg)


def _info(msg: str) -> None:
    messagebox.showinfo("PDM (Macro SolidWorks)", msg)


def _is_callable_or_value(x):
    try:
        return x() if callable(x) else x
    except Exception:
        try:
            return x()
        except Exception:
            return x


def _detect_doc_type(sw_doc: Any, path: str) -> Optional[DocType]:
    # Prefer file extension
    if path:
        ext = Path(path).suffix.lower()
        if ext == ".sldprt":
            return "PART"
        if ext == ".sldasm":
            return "ASSY"
        return None
    # Fallback: ModelDoc2.GetType -> swDocumentTypes_e (1=part,2=assy,3=drw)
    try:
        t = getattr(sw_doc, "GetType", None)
        v = _is_callable_or_value(t)
        if v == 1:
            return "PART"
        if v == 2:
            return "ASSY"
    except Exception:
        pass
    return None


def _best_sw_prop_for(cfg: AppConfig, pdm_field: str, default: str) -> str:
    try:
        for it in (cfg.solidworks.property_mappings or []):
            if str(it.get("pdm_field", "")).strip() == pdm_field:
                sp = str(it.get("sw_prop", "")).strip()
                if sp:
                    return sp
    except Exception:
        pass
    try:
        sp = (cfg.solidworks.property_map or {}).get(pdm_field, "")
        if sp:
            return sp
    except Exception:
        pass
    return default


def _core_props_from_doc(doc: Document) -> Dict[str, str]:
    # Core fisso richiesto: code, revision, state, doc_type, mmm, gggg, vvv
    return {
        "code": doc.code,
        "revision": f"{doc.revision:02d}",
        "state": doc.state,
        "doc_type": doc.doc_type,
        "mmm": doc.mmm,
        "gggg": doc.gggg,
        "vvv": doc.vvv or "",
    }


def _map_props_to_sw(cfg: AppConfig, core: Dict[str, str]) -> Dict[str, str]:
    out: Dict[str, str] = {}
    mappings = []
    try:
        mappings = list(cfg.solidworks.property_mappings or [])
    except Exception:
        mappings = []
    if not mappings:
        try:
            mappings = [{"pdm_field": k, "sw_prop": v} for k, v in (cfg.solidworks.property_map or {}).items()]
        except Exception:
            mappings = []
    for it in mappings:
        pf = str(it.get("pdm_field", "")).strip()
        sp = str(it.get("sw_prop", "")).strip()
        if not pf or not sp:
            continue
        if pf in core:
            out[sp] = str(core[pf])
    return out


def _read_sw_custom(cfg: AppConfig, sw_doc: Any) -> Dict[str, str]:
    props = get_custom_properties(sw_doc) or {}
    out: Dict[str, str] = {}
    desc_prop = str(getattr(cfg.solidworks, "description_prop", "DESCRIZIONE") or "DESCRIZIONE")
    if desc_prop:
        out["description"] = str(props.get(desc_prop, "") or "")
    for p in (getattr(cfg.solidworks, "read_properties", []) or []):
        p = str(p or "").strip()
        if not p:
            continue
        out[p] = str(props.get(p, "") or "")
    return out



def _code_from_path(p: str) -> str:
    """Estrae il codice base da un percorso file SolidWorks.
    Gestisce suffix di INREV ( _R00__INREV ) e REV ( _R00 ).
    """
    try:
        stem = Path(p).stem
    except Exception:
        stem = (p or "")
    stem = (stem or "").strip()
    if not stem:
        return ""
    # INREV: CODE_R00__INREV
    stem = _re.sub(r"_R\d{2}__INREV$", "", stem, flags=_re.I)
    # REV: CODE_R00
    stem = _re.sub(r"_R\d{2}$", "", stem, flags=_re.I)
    return stem


DOC_LOCK_TTL_SECONDS = 20 * 60


def _load_sw_context(ns) -> dict:
    """Carica contesto SolidWorks.
    Preferisce file JSON (sw_context_file). Accetta anche JSON stringa (sw_context).
    In caso di JSON non valido (es. backslash non escapate), tenta un fallback regex.
    """
    s = ""
    # 1) File
    fpath = getattr(ns, "sw_context_file", "") or ""
    if fpath:
        try:
            p = Path(fpath)
            if p.exists():
                s = p.read_text(encoding="utf-8", errors="replace")
        except Exception:
            s = ""
    # 2) Stringa
    if not s:
        s = getattr(ns, "sw_context", "") or ""

    s = (s or "").strip()
    if not s:
        return {}

    # Prova JSON valido
    try:
        return json.loads(s)
    except Exception:
        pass

    # Fallback: estrai active_doc_path con regex e de-escapa \ ->     try:
        m = _re.search(r'"active_doc_path"\s*:\s*"([^"]*)"', s)
        if m:
            val = m.group(1)
            val = val.replace('\\\\', '\\')  # JSON double backslash -> single
            return {"active_doc_path": val}
    except Exception:
        pass

    return {}

if ctk is not None:

    class MacroUI(ctk.CTk):
        """UI macro lanciata da SolidWorks.
        WORKSPACE bloccata (id passato dal bootstrap).
        """

        def __init__(self, pdm_root: Path, ws_id: str, sw_context: dict):
            super().__init__()
            self.title("PDM — Macro SolidWorks")
            self.geometry("980x680")
            self.minsize(920, 620)

            self.pdm_root = Path(pdm_root)
            self.ws_id = (ws_id or "").strip()
            self.sw_context = sw_context or {}
            self.sw_pid = None
            try:
                _pid = self.sw_context.get("sw_pid")
                if _pid is not None:
                    self.sw_pid = int(_pid)
            except Exception:
                self.sw_pid = None
            self.log_file = str((self.pdm_root / "SW_CACHE" / self.ws_id / "payload" / "payload.log"))

            # Managers
            self.ws_mgr = WorkspaceManager(self.pdm_root / "WORKSPACES")
            self.ws = self.ws_mgr.get(self.ws_id)
            if not self.ws:
                raise ValueError(f"Workspace non trovata: {self.ws_id}")
            self.workflow_log_file = str(self.ws_mgr.workspace_dir(self.ws_id) / "LOGS" / "workflow.log")

            self.cfg_mgr = ConfigManager(self.ws_mgr.config_path(self.ws_id))
            self.cfg = self.cfg_mgr.load()

            self.store = Store(self.ws_mgr.db_path(self.ws_id))
            pdm_user_hint = str(self.sw_context.get("pdm_user") or self.sw_context.get("pdm_username") or "").strip()
            self.session = resolve_session_context(pdm_user_hint=pdm_user_hint)
            self.lock_ttl_seconds = DOC_LOCK_TTL_SECONDS

            # SW
            self.sw = None

            # Vars UI
            self.active_doc_path = (self.sw_context.get("active_doc_path") or "").strip()
            self.code_var = tk.StringVar(value=_code_from_path(self.active_doc_path))
            self.use_vvv_var = tk.BooleanVar(value=bool(self.cfg.code.include_vvv_by_default))

            self.mmm_var = tk.StringVar(value="")
            self.gggg_var = tk.StringVar(value="")
            self.vvv_var = tk.StringVar(value=(self.cfg.code.vvv_presets[0] if (self.cfg.code.vvv_presets or []) else ""))
            self.desc_var = tk.StringVar(value="")

            self.doc_type = self._detect_doc_type_from_path(self.active_doc_path)
            
            # Nuove variabili per tab codifica allineata a PDM
            self.doc_type_var = tk.StringVar(value=self.doc_type if self.doc_type in ("MACHINE", "GROUP", "PART", "ASSY") else "PART")
            self.file_mode_var = tk.StringVar(value="code_only")  # code_only, model, model_drw
            self.link_file_var = tk.StringVar(value="")
            self.link_auto_drw_var = tk.BooleanVar(value=True)

            self._build_ui()
            _handler = getattr(self, "_on_close", None)
            if not callable(_handler):
                _handler = self.destroy
            self.protocol("WM_DELETE_WINDOW", _handler)
            self.after(50, self._refresh_from_active_doc)
            self._log_activity("MACRO_START", message=f"Macro avviata | source={self.session.get('source','UNKNOWN')}")

        # ---------------- UI base ----------------
        def _build_ui(self):
            top = ctk.CTkFrame(self)
            top.pack(fill="x", padx=10, pady=10)

            ws_name = getattr(self.ws, "name", "") or ""
            ws_desc = getattr(self.ws, "description", "") or ""
            title = f"WORKSPACE: {self.ws_id}"
            if ws_name or ws_desc:
                title += f" — {ws_name} | {ws_desc}"

            ctk.CTkLabel(top, text=title, font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=8, pady=(6, 2))
            ctk.CTkLabel(
                top,
                text=f"Utente: {self.session.get('display_name', 'unknown')}",
                font=ctk.CTkFont(size=12, weight="bold"),
            ).pack(anchor="w", padx=8, pady=(0, 2))

            self.lbl_doc = ctk.CTkLabel(top, text=self._doc_label_text(), font=ctk.CTkFont(size=12))
            self.lbl_doc.pack(anchor="w", padx=8, pady=(0, 8))

            # Tab view
            self.tabs = ctk.CTkTabview(self)
            self.tabs.pack(fill="both", expand=True, padx=10, pady=(0, 10))
            self.tab_cod = self.tabs.add("Codifica")
            self.tab_wf = self.tabs.add("Workflow")

            self._ui_codifica()
            self._ui_workflow()

        def _doc_label_text(self) -> str:
            p = (self.active_doc_path or "").strip()
            if p:
                return f"Documento SolidWorks attivo: {p}"
            return "Documento SolidWorks attivo: (nessuno) — apri un PART o ASSY."


        def _log_activity(
            self,
            action: str,
            code: str = "",
            status: str = "OK",
            message: str = "",
            details: dict | None = None,
        ) -> None:
            try:
                self.store.add_activity(
                    workspace_id=self.ws_id,
                    session_id=str(self.session.get("session_id", "")),
                    user_id=str(self.session.get("user_id", "")),
                    user_display=str(self.session.get("display_name", "")),
                    host=str(self.session.get("host", "")),
                    action=action,
                    code=code,
                    status=status,
                    message=message,
                    details=details or {},
                )
            except Exception:
                pass

        def _acquire_doc_lock(self, code: str, action: str) -> bool:
            ok, lock_status, holder = self.store.acquire_document_lock(
                code=code,
                owner_session=str(self.session.get("session_id", "")),
                owner_user=str(self.session.get("display_name", "")),
                owner_host=str(self.session.get("host", "")),
                ttl_seconds=self.lock_ttl_seconds,
            )
            if ok:
                return True
            who = str(holder.get("owner_user", "") or holder.get("owner_session", "altro utente"))
            host = str(holder.get("owner_host", "") or "")
            msg = f"Documento {code} bloccato da {who}" + (f" su {host}" if host else "") + "."
            self._log_activity(action=action, code=code, status="LOCKED", message=lock_status, details={"holder": holder})
            messagebox.showwarning("PDM (Macro SolidWorks)", msg + "\nRiprova tra poco.")
            return False

        def _release_doc_lock(self, code: str) -> None:
            try:
                self.store.release_document_lock(code=code, owner_session=str(self.session.get("session_id", "")))
            except Exception:
                pass


        def _ensure_sw(self) -> Any:
            """Restituisce l'istanza SolidWorks corretta (quella che ha il documento attivo).
            sw_api.get_solidworks_app() ritorna (sw_app, SWResult), quindi qui facciamo unwrap e
            passiamo preferenze per selezionare l'istanza giusta.
            """
            if self.sw is None:
                prefer_doc = (self.active_doc_path or "").strip() or None
                prefer_pid = self.sw_pid
                sw, res = get_solidworks_app(
                    visible=True,
                    timeout_s=10.0,
                    allow_launch=False,
                    prefer_pid=prefer_pid,
                    prefer_doc_path=prefer_doc,
                )
                if not sw:
                    raise RuntimeError(f"SolidWorks non disponibile: {getattr(res, 'message', res)}")
                self.sw = sw
            return self.sw

        def _active_doc(self) -> Any:
            sw = self._ensure_sw()
            try:
                return sw.ActiveDoc
            except Exception:
                return None

        def _detect_doc_type_from_path(self, p: str) -> str:
            """Ritorna tipo documento in formato compatibile col PDM: PART/ASSY."""
            ext = (Path(p).suffix or "").lower()
            if ext == ".sldprt":
                return "PART"
            if ext == ".sldasm":
                return "ASSY"
            # Se lanciato da DRW o sconosciuto, default PART (non blocca UI)
            return "PART"

        def _doc_type_label(self, t: str) -> str:
            """Etichetta breve per UI."""
            tt = (t or "").upper()
            if tt == "PART":
                return "PRT"
            if tt == "ASSY":
                return "ASSY"
            return tt

        def _normalize_segment(self, seg: str, value: str) -> str:
            """Normalizza MMM/GGGG/VVV usando le regole configurate in code.segments."""
            v = (value or "").strip()
            try:
                rule = (self.cfg.code.segments or {}).get(seg)
                if rule is not None:
                    return rule.normalize_value(v)
            except Exception:
                pass
            return v.upper()

        def _refresh_from_active_doc(self):
            # Rileva path reale dal SW (se possibile)
            try:
                sw = self._ensure_sw()
                doc = sw.ActiveDoc
                if doc is not None:
                    pn = ""
                    try:
                        pn = doc.GetPathName()
                    except Exception:
                        pn = ""
                    if pn:
                        self.active_doc_path = pn
            except Exception:
                pass

            self.doc_type = self._detect_doc_type_from_path(self.active_doc_path)
            self.code_var.set(_code_from_path(self.active_doc_path))

            try:
                sw = self._ensure_sw()
                doc = sw.ActiveDoc
                if doc is not None:
                    props = _read_sw_custom(self.cfg, doc)
                    desc = str(props.get("description", "") or "").strip()
                    if desc:
                        self.desc_var.set(desc)
            except Exception:
                pass

            self.lbl_doc.configure(text=self._doc_label_text())
            self._refresh_machines()
            self._refresh_preview()

        # ---------------- CODIFICA ----------------
        def _ui_codifica(self):
            tab = self.tab_cod

            frm = ctk.CTkFrame(tab)
            frm.pack(fill="both", expand=True, padx=12, pady=12)

            # --- TIPO DOCUMENTO ---
            type_frame = ctk.CTkFrame(frm)
            type_frame.pack(fill="x", pady=(0, 10))
            ctk.CTkLabel(type_frame, text="TIPO DOCUMENTO", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=8, pady=6)
            
            radio_frame = ctk.CTkFrame(type_frame, fg_color="transparent")
            radio_frame.pack(fill="x", padx=8, pady=4)
            
            ctk.CTkRadioButton(radio_frame, text="Macchina (MMM-V####) → crea ASM", variable=self.doc_type_var, value="MACHINE", command=self._on_doc_type_change).pack(anchor="w", pady=2)
            ctk.CTkRadioButton(radio_frame, text="Gruppo (MMM_GGGG-V####) → crea ASM", variable=self.doc_type_var, value="GROUP", command=self._on_doc_type_change).pack(anchor="w", pady=2)
            ctk.CTkRadioButton(radio_frame, text="Parte (MMM_GGGG-0001) → crea PRT", variable=self.doc_type_var, value="PART", command=self._on_doc_type_change).pack(anchor="w", pady=2)
            ctk.CTkRadioButton(radio_frame, text="Assieme (MMM_GGGG-9999) → crea ASM", variable=self.doc_type_var, value="ASSY", command=self._on_doc_type_change).pack(anchor="w", pady=2)

            # --- PARAMETRI ---
            params_frame = ctk.CTkFrame(frm)
            params_frame.pack(fill="x", pady=(0, 10))
            
            params_grid = ctk.CTkFrame(params_frame, fg_color="transparent")
            params_grid.pack(fill="x", padx=8, pady=8)
            
            ctk.CTkLabel(params_grid, text="MMM").grid(row=0, column=0, padx=6, pady=6, sticky="w")
            self.cb_mmm = ctk.CTkComboBox(params_grid, variable=self.mmm_var, values=[], command=lambda _: self._on_mmm_changed(), width=140)
            self.cb_mmm.grid(row=0, column=1, padx=6, pady=6, sticky="w")
            
            ctk.CTkLabel(params_grid, text="GGGG").grid(row=0, column=2, padx=6, pady=6, sticky="w")
            self.cb_gggg = ctk.CTkComboBox(params_grid, variable=self.gggg_var, values=[], command=lambda _: self._refresh_preview(), width=140)
            self.cb_gggg.grid(row=0, column=3, padx=6, pady=6, sticky="w")
            
            self.chk_use_vvv = ctk.CTkCheckBox(params_grid, text="Variante", variable=self.use_vvv_var, command=self._refresh_preview)
            self.chk_use_vvv.grid(row=0, column=4, padx=6, pady=6, sticky="w")
            
            vvv_vals = list(self.cfg.code.vvv_presets or ["V01"])
            self.cb_vvv = ctk.CTkComboBox(params_grid, variable=self.vvv_var, values=vvv_vals, command=lambda _: self._refresh_preview(), width=120)
            self.cb_vvv.grid(row=0, column=5, padx=6, pady=6, sticky="w")
            
            ctk.CTkLabel(params_grid, text="Descrizione").grid(row=1, column=0, padx=6, pady=6, sticky="w")
            self.ent_desc = ctk.CTkEntry(params_grid, textvariable=self.desc_var, width=500)
            self.ent_desc.grid(row=1, column=1, columnspan=5, padx=6, pady=6, sticky="ew")
            
            params_grid.grid_columnconfigure(5, weight=1)

            # --- FILE ESISTENTE (opzionale) ---
            link_frame = ctk.CTkFrame(params_frame, fg_color="transparent")
            link_frame.pack(fill="x", padx=8, pady=(0, 8))
            ctk.CTkLabel(link_frame, text="File esistente (opz)").pack(side="left", padx=6)
            self.ent_link = ctk.CTkEntry(link_frame, textvariable=self.link_file_var, width=300)
            self.ent_link.pack(side="left", padx=6)
            ctk.CTkButton(link_frame, text="Sfoglia", width=90, command=self._browse_link_file).pack(side="left", padx=4)
            ctk.CTkButton(link_frame, text="Pulisci", width=80, command=self._clear_link_file).pack(side="left", padx=4)
            ctk.CTkCheckBox(link_frame, text="Importa anche DRW", variable=self.link_auto_drw_var).pack(side="left", padx=12)

            # --- CREAZIONE FILE SOLIDWORKS ---
            file_frame = ctk.CTkFrame(frm)
            file_frame.pack(fill="x", pady=(0, 10))
            ctk.CTkLabel(file_frame, text="CREAZIONE FILE SOLIDWORKS", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=8, pady=6)
            
            file_radio_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
            file_radio_frame.pack(fill="x", padx=8, pady=4)
            
            ctk.CTkRadioButton(file_radio_frame, text="Solo codice (no file)", variable=self.file_mode_var, value="code_only").pack(anchor="w", pady=2)
            ctk.CTkRadioButton(file_radio_frame, text="Modello (PRT/ASM automatico)", variable=self.file_mode_var, value="model").pack(anchor="w", pady=2)
            ctk.CTkRadioButton(file_radio_frame, text="Modello + Disegno", variable=self.file_mode_var, value="model_drw").pack(anchor="w", pady=2)

            # --- AZIONI ---
            actions_frame = ctk.CTkFrame(frm)
            actions_frame.pack(fill="x", pady=(10, 0))
            
            left_actions = ctk.CTkFrame(actions_frame, fg_color="transparent")
            left_actions.pack(side="left", padx=8, pady=8)
            ctk.CTkButton(left_actions, text="PROSSIMO CODICE", width=160, command=self._show_next_code).pack(side="left", padx=6)
            
            self.lbl_preview = ctk.CTkLabel(left_actions, text="", font=ctk.CTkFont(size=16, weight="bold"), text_color="#2E7D32")
            self.lbl_preview.pack(side="left", padx=12)
            
            right_actions = ctk.CTkFrame(actions_frame, fg_color="transparent")
            right_actions.pack(side="right", padx=8, pady=8)
            ctk.CTkButton(right_actions, text="GENERA", width=180, height=40, font=ctk.CTkFont(size=16, weight="bold"), fg_color="#27AE60", hover_color="#229954", command=self._generate_document).pack()

            self._on_doc_type_change()
            self._refresh_preview()

        def _refresh_machines(self):
            try:
                machines = self.store.list_machines()
            except Exception:
                machines = []
            mmm_codes = [m[0] for m in machines]  # (code, desc)
            if not mmm_codes:
                mmm_codes = [""]

            # seleziona primo se vuoto
            if not self.mmm_var.get() and mmm_codes and mmm_codes[0]:
                self.mmm_var.set(mmm_codes[0])

            try:
                self.cb_mmm.configure(values=mmm_codes)
            except Exception:
                pass

            self._on_mmm_changed()

        def _on_mmm_changed(self):
            mmm = (self.mmm_var.get() or "").strip()
            try:
                groups = self.store.list_groups(mmm)
            except Exception:
                groups = []
            g_codes = [g[0] for g in groups]  # (code, desc, mmm)
            if not g_codes:
                g_codes = [""]

            if (not self.gggg_var.get()) and g_codes and g_codes[0]:
                self.gggg_var.set(g_codes[0])

            try:
                self.cb_gggg.configure(values=g_codes)
            except Exception:
                pass

            self._refresh_preview()

        def _on_doc_type_change(self):
            """Abilita/disabilita controlli in base al tipo documento selezionato."""
            doc_type = self.doc_type_var.get()
            
            if doc_type == "MACHINE":
                # Disabilita GGGG e variante
                self.cb_gggg.configure(state="disabled")
                self.chk_use_vvv.configure(state="disabled")
                self.cb_vvv.configure(state="disabled")
            elif doc_type == "GROUP":
                # Abilita GGGG, disabilita variante
                self.cb_gggg.configure(state="normal")
                self.chk_use_vvv.configure(state="disabled")
                self.cb_vvv.configure(state="disabled")
            else:  # PART/ASSY
                # Abilita tutto
                self.cb_gggg.configure(state="normal")
                self.chk_use_vvv.configure(state="normal")
                vvv_state = "normal" if self.use_vvv_var.get() else "disabled"
                try:
                    self.cb_vvv.configure(state=vvv_state)
                except Exception:
                    pass
            
            self._refresh_preview()

        def _browse_link_file(self):
            """Apre dialog per selezionare un file SolidWorks esistente."""
            from tkinter import filedialog
            path = filedialog.askopenfilename(
                title='Seleziona file SolidWorks',
                filetypes=[('SolidWorks', '*.sldprt *.sldasm *.slddrw'), ('Tutti i file', '*.*')]
            )
            if path:
                self.link_file_var.set(path)

        def _clear_link_file(self):
            """Pulisce il campo file esistente."""
            self.link_file_var.set('')

        def _require_desc_upper(self, value: str, what: str = "campo") -> str | None:
            """Valida e normalizza descrizione richiesta (maiuscolo)."""
            v = (value or "").strip()
            if not v:
                messagebox.showwarning("PDM (Macro SolidWorks)", f"Il {what} è obbligatorio.")
                return None
            return v.upper()

        def _show_next_code(self):
            """Calcola e mostra il prossimo codice disponibile."""
            self._refresh_preview()

        def _refresh_preview(self):
            """Calcola e mostra il prossimo codice in base al tipo documento e parametri selezionati."""
            doc_type = self.doc_type_var.get()
            mmm = (self.mmm_var.get() or "").strip()

            try:
                if doc_type == "MACHINE":
                    if not mmm:
                        self.lbl_preview.configure(text="")
                        return
                    row = self.store.conn.execute(
                        "SELECT next_ver FROM ver_counters WHERE mmm=? AND gggg='' AND doc_type='MACHINE';",
                        (mmm,)
                    ).fetchone()
                    seq = int(row["next_ver"]) if row else 1
                    from pdm_sw.codegen import build_machine_code
                    code = build_machine_code(self.cfg, mmm, seq)
                    self.lbl_preview.configure(text=f"Prossimo: {code}")
                
                elif doc_type == "GROUP":
                    gggg = (self.gggg_var.get() or "").strip()
                    if not mmm or not gggg:
                        self.lbl_preview.configure(text="")
                        return
                    row = self.store.conn.execute(
                        "SELECT next_ver FROM ver_counters WHERE mmm=? AND gggg=? AND doc_type='GROUP';",
                        (mmm, gggg)
                    ).fetchone()
                    seq = int(row["next_ver"]) if row else 1
                    from pdm_sw.codegen import build_group_code
                    code = build_group_code(self.cfg, mmm, gggg, seq)
                    self.lbl_preview.configure(text=f"Prossimo: {code}")
                
                else:  # PART/ASSY
                    gggg = (self.gggg_var.get() or "").strip()
                    if not mmm or not gggg:
                        self.lbl_preview.configure(text="")
                        return
                    vvv = (self.vvv_var.get() or "").strip().upper() if self.use_vvv_var.get() else ""
                    seq = self.store.peek_seq(mmm, gggg, vvv, doc_type)
                    code = build_code(self.cfg, mmm, gggg, seq, vvv=vvv, force_vvv=self.use_vvv_var.get())
                    self.lbl_preview.configure(text=f"Prossimo: {code}")
                    
            except Exception as e:
                self.lbl_preview.configure(text=f"Errore: {e}")

        def _generate_document(self):
            """Genera documento con supporto MACHINE, GROUP, PART, ASSY + import file."""
            doc_type = self.doc_type_var.get()
            file_mode = self.file_mode_var.get()
            mmm = (self.mmm_var.get() or "").strip()
            
            # Validazione parametri base
            if not mmm:
                messagebox.showwarning("PDM (Macro SolidWorks)", "Seleziona MMM.")
                return
            
            desc = self._require_desc_upper(self.desc_var.get(), what="descrizione")
            if desc is None:
                return
            self.desc_var.set(desc)
            
            # Archivio necessario per creazione file
            archive_root = Path(self.cfg.solidworks.archive_root or "").expanduser()
            if not str(archive_root).strip() and file_mode != "code_only":
                messagebox.showwarning("PDM (Macro SolidWorks)", "Archivio (root) non configurato nel PDM (tab SolidWorks).")
                return
            
            # Validazione specifica per tipo e allocazione codice
            gggg = ""
            vvv = ""
            code = ""
            seq = 0
            action_log = ""
            wip_path = ""
            wip_drw_path = ""
            
            try:
                from pdm_sw.codegen import build_machine_code, build_group_code
                
                if doc_type == "MACHINE":
                    # MACHINE: solo MMM necessario
                    seq = self.store.allocate_ver_seq(mmm, "", "MACHINE")
                    code = build_machine_code(self.cfg, mmm, seq)
                    
                    if str(archive_root).strip():
                        from pdm_sw.archive import archive_dirs_for_machine
                        wip, rel, inrev, rev = archive_dirs_for_machine(archive_root, mmm)
                        wip_path = str(model_path(wip, code, "MACHINE"))
                        if file_mode == "model_drw":
                            wip_drw_path = str(drw_path(wip, code))
                    action_log = "CREATE_MACHINE"
                    
                elif doc_type == "GROUP":
                    # GROUP: MMM + GGGG necessari
                    gggg = (self.gggg_var.get() or "").strip()
                    if not gggg:
                        messagebox.showwarning("PDM (Macro SolidWorks)", "Seleziona GGGG.")
                        return
                    
                    seq = self.store.allocate_ver_seq(mmm, gggg, "GROUP")
                    code = build_group_code(self.cfg, mmm, gggg, seq)
                    
                    if str(archive_root).strip():
                        from pdm_sw.archive import archive_dirs_for_group
                        wip, rel, inrev, rev = archive_dirs_for_group(archive_root, mmm, gggg)
                        wip_path = str(model_path(wip, code, "GROUP"))
                        if file_mode == "model_drw":
                            wip_drw_path = str(drw_path(wip, code))
                    action_log = "CREATE_GROUP"
                    
                else:  # PART or ASSY
                    # PART/ASSY: MMM + GGGG + opzionale variante
                    gggg = (self.gggg_var.get() or "").strip()
                    if not gggg:
                        messagebox.showwarning("PDM (Macro SolidWorks)", "Seleziona GGGG.")
                        return
                    
                    vvv = (self.vvv_var.get() or "").strip() if self.use_vvv_var.get() else ""
                    vvv_norm = self._normalize_segment("VVV", vvv) if vvv else ""
                    seq = self.store.allocate_seq(mmm, gggg, vvv_norm, doc_type)
                    vvv = vvv_norm
                    code = build_code(self.cfg, mmm, gggg, seq, vvv=vvv, force_vvv=bool(vvv))
                    
                    if str(archive_root).strip():
                        wip, rel, inrev, rev = archive_dirs(archive_root, mmm, gggg)
                        wip_path = str(model_path(wip, code, doc_type))
                        if file_mode == "model_drw":
                            wip_drw_path = str(drw_path(wip, code))
                    action_log = "CREATE_CODE"
                
                # Crea record documento in DB
                new_doc = Document(
                    id=0,
                    code=code,
                    doc_type=doc_type,
                    mmm=mmm,
                    gggg=gggg,
                    seq=seq,
                    vvv=vvv,
                    revision=0,
                    state="WIP",
                    obs_prev_state="",
                    description=desc,
                    file_wip_path=wip_path,
                    file_rel_path="",
                    file_inrev_path="",
                    file_wip_drw_path=wip_drw_path,
                    file_rel_drw_path="",
                    file_inrev_drw_path="",
                    created_at="",
                    updated_at="",
                )
                
                # Gestione file: import esistente o creazione da template o SaveAs documento attivo
                link_file = self.link_file_var.get().strip()
                sw = self._ensure_sw()
                sw_doc_active = None
                
                try:
                    sw_doc_active = self._active_doc()
                except Exception:
                    sw_doc_active = None
                
                if link_file and doc_type in ("PART", "ASSY"):
                    # Import file esistente
                    self._import_linked_file_to_wip(new_doc, link_file, file_mode == "model_drw")
                    
                elif file_mode in ("model", "model_drw") and sw_doc_active is not None:
                    # SaveAs documento SolidWorks attivo
                    if not wip_path:
                        messagebox.showwarning("PDM (Macro SolidWorks)", "Path WIP non disponibile (archivio non configurato).")
                        return
                    
                    # Crea cartelle
                    Path(wip_path).parent.mkdir(parents=True, exist_ok=True)
                    
                    # SaveAs
                    res = save_as_doc(sw_doc_active, wip_path)
                    if not getattr(res, "ok", False):
                        msg = getattr(res, "message", "SaveAs fallito.")
                        messagebox.showwarning("PDM (Macro SolidWorks)", f"Errore SaveAs: {msg}")
                        return
                    
                    # Aggiorna path attivo
                    self.active_doc_path = wip_path
                    
                    # Scrivi props core su SW
                    try:
                        self._write_core_props_to_sw(new_doc, sw_doc_active)
                    except Exception:
                        pass
                    
                    # Descrizione su SW
                    dp = (getattr(self.cfg.solidworks, 'description_prop', 'DESCRIZIONE') or 'DESCRIZIONE').strip()
                    if dp and desc:
                        set_custom_properties(sw_doc_active, {dp: desc})
                
                elif file_mode in ("model", "model_drw"):
                    # Creazione da template
                    if not wip_path:
                        messagebox.showwarning("PDM (Macro SolidWorks)", "Path WIP non disponibile (archivio non configurato).")
                        return
                    
                    tpl_model = ""
                    if doc_type in ("PART", "MACHINE"):
                        tpl_model = self.cfg.solidworks.template_part
                    elif doc_type in ("ASSY", "GROUP"):
                        tpl_model = self.cfg.solidworks.template_assembly
                    
                    if not tpl_model or not Path(tpl_model).exists():
                        messagebox.showwarning("PDM (Macro SolidWorks)", f"Template non configurato o non trovato per {doc_type}.")
                        return
                    
                    # Crea file modello
                    res_model = create_model_file(sw, tpl_model, wip_path, props={})
                    if not getattr(res_model, "ok", False):
                        msg = getattr(res_model, "message", "Creazione modello fallita.")
                        messagebox.showwarning("PDM (Macro SolidWorks)", f"Errore creazione modello: {msg}")
                        return
                    
                    # Crea disegno se richiesto
                    if file_mode == "model_drw" and wip_drw_path:
                        tpl_drw = self.cfg.solidworks.template_drawing
                        if tpl_drw and Path(tpl_drw).exists():
                            create_drawing_file(sw, tpl_drw, wip_drw_path, props={})
                
                # Inserisci documento in DB
                self.store.add_document(new_doc)
                
                # Log attività
                self._log_activity(action_log, code=code, status="OK", message=f"Creato {doc_type}")
                
                # Refresh UI
                self.code_var.set(code)
                self.lbl_doc.configure(text=self._doc_label_text())
                self._refresh_preview()
                self._refresh_wf_state()
                
                messagebox.showinfo("PDM (Macro SolidWorks)", f"Documento creato: {code}")
                
            except Exception as e:
                self._log_activity(action_log if action_log else "CREATE_ERROR", 
                                 code=code if code else "", status="ERROR", message=str(e))
                messagebox.showwarning("PDM (Macro SolidWorks)", f"Errore creazione documento: {e}")

        def _import_linked_file_to_wip(self, doc: Document, src_path: str, auto_drw: bool):
            """Importa file esistente e opzionalmente DRW nella cartella WIP."""
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
                    raise ValueError('Hai selezionato un DRW ma non trovo il modello (.sldprt/.sldasm) con stesso nome.')
            
            # check coerenza tipo
            if src_model.suffix.lower() == '.sldprt' and doc.doc_type not in ('PART', 'MACHINE'):
                raise ValueError("Il file selezionato è una PARTE (.sldprt) ma il tipo scelto è ASSY/GROUP.")
            if src_model.suffix.lower() == '.sldasm' and doc.doc_type not in ('ASSY', 'GROUP'):
                raise ValueError("Il file selezionato è un ASSIEME (.sldasm) ma il tipo scelto è PART/MACHINE.")

            if not doc.file_wip_path:
                raise ValueError("Path WIP non disponibile per il documento.")
            
            dst_model = Path(doc.file_wip_path)
            if dst_model.exists():
                raise ValueError("Esiste già un file modello in archivio con questo codice (WIP).")
            
            dst_model.parent.mkdir(parents=True, exist_ok=True)
            safe_copy(src_model, dst_model)
            set_readonly(dst_model, readonly=False)
            self.store.update_document(doc.code, file_wip_path=str(dst_model))

            # DRW: stesso nome del modello, stessa cartella
            if auto_drw:
                src_drw = src_model.with_suffix('.slddrw')
                if src_drw.exists() and doc.file_wip_drw_path:
                    dst_drw = Path(doc.file_wip_drw_path)
                    if not dst_drw.exists():
                        safe_copy(src_drw, dst_drw)
                        set_readonly(dst_drw, readonly=False)
                    self.store.update_document(doc.code, file_wip_drw_path=str(dst_drw))

        def _write_core_props_to_sw(self, docrec: Document, sw_doc: Any):
            # mapping PDM->SW: usa solo quelle marcate 'enabled'
            props = {
                "code": docrec.code,
                "revision": f"{int(docrec.revision):02d}",
                "state": str(docrec.state),
                "doc_type": str(docrec.doc_type),
                "mmm": docrec.mmm,
                "gggg": docrec.gggg,
                "vvv": docrec.vvv,
            }
            out = {}
            for m in (self.cfg.solidworks.property_mappings or []):
                if isinstance(m, dict):
                    if not bool(m.get("enabled", True)):
                        continue
                    pdm_name = str(m.get("pdm_field", "") or "").strip()
                    sw_name = str(m.get("sw_prop", "") or "").strip()
                else:
                    if not getattr(m, "enabled", True):
                        continue
                    pdm_name = str(getattr(m, "pdm_name", "") or "").strip()
                    sw_name = str(getattr(m, "sw_name", "") or "").strip()
                if not pdm_name or not sw_name:
                    continue
                if pdm_name in props:
                    out[sw_name] = str(props[pdm_name])
            if not out:
                try:
                    for pdm_name, sw_name in (self.cfg.solidworks.property_map or {}).items():
                        p = str(pdm_name or "").strip()
                        s = str(sw_name or "").strip()
                        if p and s and p in props:
                            out[s] = str(props[p])
                except Exception:
                    pass
            if out:
                set_custom_properties(sw_doc, out)

        # ---------------- WORKFLOW ----------------
        def _ui_workflow(self):
            tab = self.tab_wf

            frm = ctk.CTkFrame(tab)
            frm.pack(fill="both", expand=True, padx=12, pady=12)

            row = ctk.CTkFrame(frm)
            row.pack(fill="x", padx=10, pady=(10, 6))

            ctk.CTkLabel(row, text="Codice (da doc attivo):", width=180, anchor="w").pack(side="left")
            self.ent_code = ctk.CTkEntry(row, textvariable=self.code_var, width=320)
            self.ent_code.pack(side="left", padx=(0, 10))
            ctk.CTkButton(row, text="AGGIORNA DA DOCUMENTO ATTIVO", command=self._refresh_from_active_doc).pack(side="left")

            self.lbl_wf_info = ctk.CTkLabel(frm, text="Stato: — | Rev: —", font=ctk.CTkFont(size=13, weight="bold"))
            self.lbl_wf_info.pack(anchor="w", padx=12, pady=(10, 10))

            # Due righe di pulsanti: crea i widget nel frame corretto (niente pack(in_=...) con master diverso)
            row1 = ctk.CTkFrame(frm)
            row1.pack(fill="x", padx=10, pady=(0, 6))

            self.btn_wip_rel = ctk.CTkButton(row1, text="WIP → REL", command=lambda: self._wf_transition("WIP_REL"), width=140)
            self.btn_rel_inrev = ctk.CTkButton(row1, text="REL → IN_REV", command=lambda: self._wf_transition("REL_INREV"), width=140)
            self.btn_inrev_app = ctk.CTkButton(row1, text="IN_REV → REL (APPROVA)", command=lambda: self._wf_transition("INREV_APPROVE"), width=220)

            self.btn_wip_rel.pack(side="left", padx=6, pady=6)
            self.btn_rel_inrev.pack(side="left", padx=6, pady=6)
            self.btn_inrev_app.pack(side="left", padx=6, pady=6)

            row2 = ctk.CTkFrame(frm)
            row2.pack(fill="x", padx=10, pady=(0, 10))

            self.btn_inrev_cancel = ctk.CTkButton(row2, text="IN_REV → REL (ANNULLA)", command=lambda: self._wf_transition("INREV_CANCEL"), width=220)
            self.btn_to_obs = ctk.CTkButton(row2, text="→ OBS", command=lambda: self._wf_transition("TO_OBS"), width=120)
            self.btn_restore = ctk.CTkButton(row2, text="RIPRISTINA OBS", command=lambda: self._wf_transition("RESTORE_OBS"), width=160)

            self.btn_inrev_cancel.pack(side="left", padx=6, pady=6)
            self.btn_to_obs.pack(side="left", padx=6, pady=6)
            self.btn_restore.pack(side="left", padx=6, pady=6)
            self.after(200, self._refresh_wf_state)

        def _refresh_wf_state(self):
            code = (self.code_var.get() or "").strip()
            if not code:
                self.lbl_wf_info.configure(text="Stato: — | Rev: —")
                return
            doc = self.store.get_document(code)
            if not doc:
                self.lbl_wf_info.configure(text="Stato: (non presente in DB) | Rev: —")
                return

            self.lbl_wf_info.configure(text=f"Stato: {doc.state} | Rev: {int(doc.revision):02d}")

            # enable/disable
            st = str(doc.state)
            self.btn_wip_rel.configure(state=("normal" if st == "WIP" else "disabled"))
            self.btn_rel_inrev.configure(state=("normal" if st == "REL" else "disabled"))
            self.btn_inrev_app.configure(state=("normal" if st == "IN_REV" else "disabled"))
            self.btn_inrev_cancel.configure(state=("normal" if st == "IN_REV" else "disabled"))
            self.btn_restore.configure(state=("normal" if st == "OBS" else "disabled"))
            self.btn_to_obs.configure(state=("normal" if st != "OBS" else "disabled"))

        def _update_doc_record(self, doc: Document) -> None:
            self.store.update_document(
                doc.code,
                state=str(doc.state),
                revision=int(doc.revision),
                obs_prev_state=str(getattr(doc, "obs_prev_state", "") or ""),
                file_wip_path=str(doc.file_wip_path or ""),
                file_rel_path=str(doc.file_rel_path or ""),
                file_inrev_path=str(doc.file_inrev_path or ""),
                file_wip_drw_path=str(doc.file_wip_drw_path or ""),
                file_rel_drw_path=str(doc.file_rel_drw_path or ""),
                file_inrev_drw_path=str(doc.file_inrev_drw_path or ""),
            )

        def _workflow_note_meta(self, action: str, doc: Document) -> tuple[str, str, str]:
            """Ritorna metadati nota: (event_type, titolo, stato_destinazione_atteso)."""
            if action == "WIP_REL":
                return "RELEASE", "Release", "REL"
            if action == "REL_INREV":
                return "CREATE_REV", "Crea revisione", "IN_REV"
            if action == "INREV_APPROVE":
                return "APPROVE_REV", "Approva revisione", "REL"
            if action == "INREV_CANCEL":
                return "CANCEL_REV", "Annulla revisione", "REL"
            if action == "TO_OBS":
                return "SET_OBSOLETE", "Imposta OBS", "OBS"
            if action == "RESTORE_OBS":
                return "RESTORE_OBS", "Ripristina da OBS", (doc.obs_prev_state or "REL")
            raise ValueError("Azione non riconosciuta")

        def _prompt_workflow_note(self, code: str, event_title: str, from_state: str, to_state: str) -> str | None:
            result: dict[str, str | None] = {"note": None}

            top = ctk.CTkToplevel(self)
            top.title("Nota cambio stato")
            top.geometry("760x420")
            top.grab_set()

            ctk.CTkLabel(
                top,
                text="Compila la nota obbligatoria per il cambio stato",
                font=ctk.CTkFont(size=16, weight="bold"),
            ).pack(anchor="w", padx=12, pady=(12, 6))

            ctk.CTkLabel(
                top,
                text=f"Codice: {code} | Evento: {event_title} | {from_state} -> {to_state}",
                font=ctk.CTkFont(size=12),
            ).pack(anchor="w", padx=12, pady=(0, 8))

            box = ctk.CTkTextbox(top, height=230)
            box.pack(fill="both", expand=True, padx=12, pady=(0, 10))
            box.focus_set()

            btns = ctk.CTkFrame(top, fg_color="transparent")
            btns.pack(fill="x", padx=12, pady=(0, 12))

            def _cancel():
                result["note"] = None
                top.destroy()

            def _ok():
                note = (box.get("1.0", "end") or "").strip()
                if len(note) < 3:
                    messagebox.showwarning("PDM (Macro SolidWorks)", "Inserisci una nota di almeno 3 caratteri.")
                    return
                if len(note) > 2000:
                    messagebox.showwarning("PDM (Macro SolidWorks)", "Nota troppo lunga (massimo 2000 caratteri).")
                    return
                result["note"] = note
                top.destroy()

            ctk.CTkButton(btns, text="Annulla", width=120, command=_cancel).pack(side="right", padx=6)
            ctk.CTkButton(btns, text="Conferma", width=120, command=_ok).pack(side="right", padx=6)
            top.bind("<Escape>", lambda _e: _cancel())

            self.wait_window(top)
            return result["note"]

        def _wf_transition(self, action: str):
            code = (self.code_var.get() or "").strip()
            if not code:
                messagebox.showwarning("PDM (Macro SolidWorks)", "Nessun codice selezionato.")
                return
            doc = self.store.get_document(code)
            if not doc:
                messagebox.showwarning("PDM (Macro SolidWorks)", "Codice non presente nel DB PDM.")
                return

            archive_root = Path(self.cfg.solidworks.archive_root or "").expanduser()
            if not str(archive_root).strip():
                messagebox.showwarning("PDM (Macro SolidWorks)", "Archivio (root) non configurato nel PDM.")
                return

            try:
                event_type, event_title, expected_to = self._workflow_note_meta(action, doc)
            except Exception as e:
                messagebox.showwarning("PDM (Macro SolidWorks)", f"Workflow fallito: {e}")
                return

            from_state = str(doc.state)
            rev_before = int(doc.revision)
            note = self._prompt_workflow_note(code=doc.code, event_title=event_title, from_state=from_state, to_state=expected_to)
            if note is None:
                return
            if not self._acquire_doc_lock(doc.code, action=f"WF_{action}"):
                return

            # Chiudi documenti SolidWorks se sono tra i file che stiamo per muovere/copiare (evita WinError 32)
            try:
                try:
                    sw = self._ensure_sw()
                    candidates = [
                        doc.file_wip_path, doc.file_rel_path, doc.file_inrev_path,
                        doc.file_wip_drw_path, doc.file_rel_drw_path, doc.file_inrev_drw_path,
                    ]

                    open_docs = []
                    for p in candidates:
                        if not p:
                            continue
                        try:
                            od = sw.GetOpenDocumentByName(p) if hasattr(sw, "GetOpenDocumentByName") else None
                        except Exception:
                            od = None

                        if od is None:
                            try:
                                ad = sw.ActiveDoc
                                if ad is not None and hasattr(ad, "GetPathName"):
                                    ap = str(ad.GetPathName() or "")
                                    if os.path.normcase(os.path.normpath(ap)) == os.path.normcase(os.path.normpath(p)):
                                        od = ad
                            except Exception:
                                od = None

                        if od is not None:
                            open_docs.append((p, od))

                    if open_docs:
                        for p, od in open_docs:
                            try:
                                save_existing_doc(od)
                            except Exception:
                                pass
                            try:
                                close_doc(sw, doc=od, file_path=p)
                            except Exception:
                                pass
                except Exception:
                    pass

                # Applica transizione su file system + aggiorna DB
                if action == "WIP_REL":
                    doc, res = release_wip(doc, archive_root, log_file=self.workflow_log_file)
                elif action == "REL_INREV":
                    doc, res = create_inrev(doc, archive_root, log_file=self.workflow_log_file)
                elif action == "INREV_APPROVE":
                    doc, res = approve_inrev(doc, archive_root, log_file=self.workflow_log_file)
                elif action == "INREV_CANCEL":
                    doc, res = cancel_inrev(doc, log_file=self.workflow_log_file)
                elif action == "TO_OBS":
                    doc.obs_prev_state = str(doc.state)
                    doc, res = set_obsolete(doc, log_file=self.workflow_log_file)
                elif action == "RESTORE_OBS":
                    doc, res = restore_obsolete(doc, doc.obs_prev_state or "REL", log_file=self.workflow_log_file)
                    if res.ok:
                        doc.obs_prev_state = ""
                else:
                    raise ValueError("Azione non riconosciuta")
                if not res.ok:
                    raise RuntimeError(res.message)
                self._update_doc_record(doc)
                try:
                    self.store.add_state_note(
                        code=doc.code,
                        event_type=event_type,
                        from_state=from_state,
                        to_state=str(doc.state),
                        note=note,
                        rev_before=rev_before,
                        rev_after=int(doc.revision),
                    )
                except Exception as ne:
                    messagebox.showwarning("PDM (Macro SolidWorks)", f"Cambio stato eseguito, ma salvataggio nota fallito: {ne}")
                self._log_activity(action=f"WF_{action}", code=doc.code, status="OK", message=f"{from_state}->{doc.state}")
            except FileExistsError as e:
                self._log_activity(action=f"WF_{action}", code=doc.code, status="ERROR", message=f"File exists: {e}")
                messagebox.showwarning(
                    "PDM (Macro SolidWorks)",
                    "Workflow fallito: file destinazione gia presente.\n"
                    "Verifica storico revisioni/cartelle e riprova.\n\n"
                    f"Dettaglio: {e}",
                )
                self._refresh_wf_state()
                return
            except PermissionError as e:
                self._log_activity(action=f"WF_{action}", code=doc.code, status="ERROR", message=f"Permission: {e}")
                messagebox.showwarning(
                    "PDM (Macro SolidWorks)",
                    "Workflow fallito: file in uso o non accessibile.\n"
                    "Chiudi i file in SolidWorks/Explorer e verifica i permessi del file/cartella.\n\n"
                    f"Dettaglio: {e}",
                )
                self._refresh_wf_state()
                return
            except Exception as e:
                self._log_activity(action=f"WF_{action}", code=doc.code, status="ERROR", message=str(e))
                messagebox.showwarning("PDM (Macro SolidWorks)", f"Workflow fallito: {e}")
                self._refresh_wf_state()
                return
            finally:
                self._release_doc_lock(doc.code)

            # Riapri documento risultante (best effort)
            try:
                sw = self._ensure_sw()
                target = ""
                if str(doc.state) == "WIP":
                    target = doc.file_wip_path
                elif str(doc.state) == "REL":
                    target = doc.file_rel_path
                elif str(doc.state) == "IN_REV":
                    target = doc.file_inrev_path
                elif str(doc.state) == "OBS":
                    target = doc.best_model_path_for_state()
                if target:
                    open_doc(sw, target)
            except Exception:
                pass

            self._refresh_wf_state()
            messagebox.showinfo("PDM (Macro SolidWorks)", f"Workflow completato. Nuovo stato: {doc.state} (Rev {int(doc.revision):02d}).")

        def _on_close(self):
            try:
                self.store.release_session_locks(str(self.session.get("session_id", "")))
            except Exception:
                pass
            self._log_activity("MACRO_EXIT", status="OK", message="Macro chiusa.")
            try:
                self.store.close()
            except Exception:
                pass
            self.destroy()

else:
    class MacroUI:
        def __init__(self, *args, **kwargs):
            # Mostra errore import customtkinter (o ambiente Python errato)
            try:
                msg = f"Impossibile avviare UI macro: customtkinter non disponibile.\nDettagli: {_CTK_IMPORT_ERROR}"
                messagebox.showwarning("PDM (Macro SolidWorks)", msg)
            except Exception:
                pass
        def mainloop(self):
            return

def main(argv: Optional[list[str]] = None) -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--pdm-root", required=False, default="", help="Cartella root del PDM (dove sta app.py)")
    ap.add_argument("--workspace", required=True, help="ID interno workspace")
    ap.add_argument("--sw-context", required=False, default="", help="JSON string con contesto SolidWorks")
    ap.add_argument("--sw-context-file", required=False, default="", help="Percorso file JSON con contesto SolidWorks (consigliato)")
    ap.add_argument("--log-file", required=False, default="", help="Percorso file log (payload)")
    ns, unknown = ap.parse_known_args(argv)
    if not getattr(ns, "log_file", ""):
        ns.log_file = str(Path(__file__).with_name("payload.log"))
    _log_line(ns.log_file, f'Payload start | exe={__import__("sys").executable} | cwd={__import__("os").getcwd()}')
    if unknown:
        _log_line(ns.log_file, 'Unknown args: ' + repr(unknown))
    pdm_root = Path(ns.pdm_root).expanduser().resolve() if ns.pdm_root else None
    if not pdm_root or not pdm_root.exists():
        here = Path(__file__).resolve()
        found = None
        for par in [here] + list(here.parents):
            if (par / "app.py").exists() and (par / "pdm_sw").is_dir() and (par / "WORKSPACES").is_dir():
                found = par
                break
        if not found:
            raise SystemExit("Impossibile determinare PDM_ROOT. Passa --pdm-root.")
        pdm_root = found
    sw_context = _load_sw_context(ns)
    _log_line(ns.log_file, 'SW context: ' + repr(sw_context))
    app = MacroUI(pdm_root=pdm_root, ws_id=ns.workspace, sw_context=sw_context)
    app.mainloop()
if __name__ == "__main__":
    import traceback, sys
    try:
        main()
    except Exception as e:
        # Best-effort: se possibile scrivi un log in payload.log (se passato come argomento)
        try:
            import argparse
            ap = argparse.ArgumentParser(add_help=False)
            ap.add_argument("--log-file", default="")
            ns, _ = ap.parse_known_args()
            _log_line(ns.log_file, "FATAL: " + repr(e))
            _log_line(ns.log_file, traceback.format_exc())
        except Exception:
            pass
        raise
