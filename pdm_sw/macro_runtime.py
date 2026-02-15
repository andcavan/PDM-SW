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

            # Riga selezioni
            row1 = ctk.CTkFrame(frm)
            row1.pack(fill="x", pady=(10, 6), padx=10)

            ctk.CTkLabel(row1, text="Macchina (MMM):", width=120, anchor="w").pack(side="left")
            self.cb_mmm = ctk.CTkComboBox(row1, variable=self.mmm_var, values=[], command=lambda _=None: self._on_mmm_changed(), width=220)
            self.cb_mmm.pack(side="left", padx=(0, 12))

            ctk.CTkLabel(row1, text="Gruppo (GGGG):", width=120, anchor="w").pack(side="left")
            self.cb_gggg = ctk.CTkComboBox(row1, variable=self.gggg_var, values=[], command=lambda _=None: self._refresh_preview(), width=220)
            self.cb_gggg.pack(side="left", padx=(0, 12))

            ctk.CTkLabel(row1, text="Tipo:", width=50, anchor="w").pack(side="left")
            self.lbl_type = ctk.CTkLabel(row1, text=self._doc_type_label(self.doc_type), width=60, anchor="w")
            self.lbl_type.pack(side="left", padx=(0, 10))

            row2 = ctk.CTkFrame(frm)
            row2.pack(fill="x", pady=(0, 8), padx=10)

            self.chk_use_vvv = ctk.CTkCheckBox(row2, text="Usa VVV", variable=self.use_vvv_var, command=self._refresh_preview)
            self.chk_use_vvv.pack(side="left", padx=(0, 10))

            ctk.CTkLabel(row2, text="Variante (VVV):", width=120, anchor="w").pack(side="left")
            vvv_vals = list(self.cfg.code.vvv_presets or [])
            if self.vvv_var.get() and self.vvv_var.get() not in vvv_vals:
                vvv_vals.insert(0, self.vvv_var.get())
            self.cb_vvv = ctk.CTkComboBox(row2, variable=self.vvv_var, values=vvv_vals, command=lambda _=None: self._refresh_preview(), width=220)
            self.cb_vvv.pack(side="left", padx=(0, 12))

            row3 = ctk.CTkFrame(frm)
            row3.pack(fill="x", pady=(0, 8), padx=10)
            ctk.CTkLabel(row3, text="Descrizione:", width=120, anchor="w").pack(side="left")
            self.ent_desc = ctk.CTkEntry(row3, textvariable=self.desc_var)
            self.ent_desc.pack(side="left", fill="x", expand=True, padx=(0, 12))

            self.lbl_preview = ctk.CTkLabel(frm, text="Prossimo codice: —", font=ctk.CTkFont(size=14, weight="bold"))
            self.lbl_preview.pack(anchor="w", padx=12, pady=(8, 8))

            btns = ctk.CTkFrame(frm)
            btns.pack(fill="x", padx=10, pady=(0, 10))

            ctk.CTkButton(btns, text="AGGIORNA PROSSIMO CODICE", command=self._refresh_preview, width=220).pack(side="left", padx=6)
            ctk.CTkButton(btns, text="CREA CODICE + SALVA IN WIP", command=self._create_code_and_save_wip, width=220).pack(side="left", padx=6)
            ctk.CTkButton(btns, text="COPIA DOCUMENTO ATTIVO", command=self._copy_active_doc_and_save_wip, width=220).pack(side="left", padx=6)

            note = "Nota: la macro crea o copia il documento attivo in WIP con nuovo codice."
            ctk.CTkLabel(frm, text=note, font=ctk.CTkFont(size=11)).pack(anchor="w", padx=12, pady=(0, 10))

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

        def _refresh_preview(self):
            self.lbl_type.configure(text=self._doc_type_label(self.doc_type))
            mmm = (self.mmm_var.get() or "").strip()
            gggg = (self.gggg_var.get() or "").strip()

            vvv = (self.vvv_var.get() or "").strip()
            use_vvv = bool(self.use_vvv_var.get())
            if not use_vvv:
                vvv = ""

            if not mmm or not gggg:
                self.lbl_preview.configure(text="Prossimo codice: — (seleziona MMM e GGGG)")
                return

            # Applica regole (upper/lower + alpha/num/len)
            try:
                mmm_n = self._normalize_segment("MMM", mmm)
                gggg_n = self._normalize_segment("GGGG", gggg)
                vvv_n = self._normalize_segment("VVV", vvv) if vvv else ""
            except Exception:
                mmm_n, gggg_n, vvv_n = mmm, gggg, vvv

            try:
                next_seq = self.store.peek_seq(mmm_n, gggg_n, vvv_n, self.doc_type)
            except Exception:
                next_seq = 1

            try:
                code = build_code(self.cfg, mmm_n, gggg_n, next_seq, vvv_n, force_vvv=bool(vvv_n))
            except Exception:
                code = f"{mmm_n}_{gggg_n}-{vvv_n}-{next_seq:04d}" if vvv_n else f"{mmm_n}_{gggg_n}-{next_seq:04d}"

            self.lbl_preview.configure(text=f"Prossimo codice: {code}")

        def _create_code_and_save_wip(self):
            # 1) Documento SolidWorks: usa prima il path fornito dal bootstrap (più affidabile di ActiveDoc)
            sw = self._ensure_sw()
            active_path = (self.active_doc_path or "").strip()

            doc = None
            try:
                doc = self._active_doc()
            except Exception:
                doc = None

            # Prova a leggere ActiveDoc subito (utile anche per documenti non ancora salvati)
            # doc già tentato sopra (ActiveDoc)
            has_path = bool(active_path)
            if (not has_path) and (doc is None):
                _warn("Apri un documento PART o ASSY in SolidWorks.")
                return
            detected = _detect_doc_type(doc, active_path)
            if not detected:
                _warn("Apri un documento PART o ASSY in SolidWorks.")
                return
            self.doc_type = detected
            try:
                self.lbl_type.configure(text=self._doc_type_label(self.doc_type))
            except Exception:
                pass

            # Recupera istanza documento (ActiveDoc può risultare None in alcune condizioni COM)
            # doc già tentato sopra (ActiveDoc)
            try:
                if doc is not None and hasattr(doc, "GetPathName"):
                    p = str(doc.GetPathName() or "")
                    if p and (not active_path):
                        self.active_doc_path = p
                        active_path = p
                    elif active_path:
                        if os.path.normcase(os.path.normpath(p)) != os.path.normcase(os.path.normpath(active_path)):
                            doc = None
            except Exception:
                doc = None

            if has_path and doc is None:
                try:
                    if hasattr(sw, "GetOpenDocumentByName"):
                        doc = sw.GetOpenDocumentByName(active_path)
                except Exception:
                    doc = None

            if has_path and doc is None:
                try:
                    doc = open_doc(sw, active_path, silent=True)
                except Exception:
                    doc = None

            if doc is None:
                _warn("Apri un documento PART o ASSY in SolidWorks.")
                return

            # 2) Selezioni
            mmm = (self.mmm_var.get() or "").strip()
            gggg = (self.gggg_var.get() or "").strip()
            vvv = (self.vvv_var.get() or "").strip()
            use_vvv = bool(self.use_vvv_var.get())
            if not use_vvv:
                vvv = ""

            if not mmm or not gggg:
                messagebox.showwarning("PDM (Macro SolidWorks)", "Seleziona MMM e GGGG.")
                return

            desc = (self.desc_var.get() or "").strip().upper()
            if not desc:
                messagebox.showwarning("PDM (Macro SolidWorks)", "Inserisci la descrizione del documento.")
                return
            self.desc_var.set(desc)

            # 3) Normalizza e valida regole
            try:
                mmm_n = self._normalize_segment("MMM", mmm)
                gggg_n = self._normalize_segment("GGGG", gggg)
                vvv_n = self._normalize_segment("VVV", vvv) if vvv else ""
            except ValueError as e:
                messagebox.showwarning("PDM (Macro SolidWorks)", f"Codifica non valida: {e}")
                return

            # 4) Alloca progressivo e crea codice
            try:
                seq = self.store.allocate_seq(mmm_n, gggg_n, vvv_n, self.doc_type)
                code = build_code(self.cfg, mmm_n, gggg_n, seq, vvv_n, force_vvv=bool(vvv_n))
            except Exception as e:
                self._log_activity("CREATE_CODE", status="ERROR", message=f"Allocazione fallita: {e}")
                messagebox.showwarning("PDM (Macro SolidWorks)", f"Errore allocazione codice: {e}")
                return

            # 5) Path archiviazione
            archive_root = Path(self.cfg.solidworks.archive_root or "").expanduser()
            if not str(archive_root).strip():
                messagebox.showwarning("PDM (Macro SolidWorks)", "Archivio (root) non configurato nel PDM (tab SolidWorks).")
                return

            wip, rel, inrev, rev = archive_dirs(archive_root, mmm_n, gggg_n)

            file_wip = model_path(wip, code, self.doc_type)
            file_wip_drw = drw_path(wip, code)

            new_doc = Document(
                id=0,
                code=code,
                doc_type=self.doc_type,
                mmm=mmm_n,
                gggg=gggg_n,
                seq=seq,
                vvv=vvv_n,
                revision=0,
                state="WIP",
                obs_prev_state="",
                description=desc,
                file_wip_path=str(file_wip),
                file_rel_path=str(model_path(rel, code, self.doc_type)),
                file_inrev_path=str(model_path(inrev, inrev_tag(code, 0), self.doc_type)),
                file_wip_drw_path=str(file_wip_drw),
                file_rel_drw_path=str(drw_path(rel, code)),
                file_inrev_drw_path=str(drw_path(inrev, inrev_tag(code, 0))),
                created_at="",
                updated_at="",
            )

            # 6) Salva in SW (SaveAs)
            try:
                # Crea cartelle
                file_wip.parent.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                messagebox.showwarning("PDM (Macro SolidWorks)", f"Errore creazione cartelle WIP: {e}")
                return

            res = save_as_doc(doc, str(file_wip))
            if not getattr(res, "ok", False):
                msg = getattr(res, "message", "SaveAs fallito.")
                det = str(getattr(res, "details", "") or "").strip()
                _log_line(self.log_file, f"SaveAs FAIL | target={file_wip} | msg={msg} | details={det}")
                messagebox.showwarning(
                    "PDM (Macro SolidWorks)",
                    f"Errore SaveAs in WIP: {msg}" + (f"\n\nDettagli: {det}" if det else "")
                )
                return
            else:
                det = str(getattr(res, "details", "") or "").strip()
                if det:
                    _log_line(self.log_file, f"SaveAs OK | target={file_wip} | details={det}")

            # 7) Inserisci in DB
            try:
                self.store.add_document(new_doc)
            except Exception as e:
                self._log_activity("CREATE_CODE", code=code, status="ERROR", message=f"DB non aggiornato: {e}")
                messagebox.showwarning("PDM (Macro SolidWorks)", f"Codice creato ma DB non aggiornato: {e}")
                return

            # 8) Scrivi core props su SW (se mappate)
            try:
                self._write_core_props_to_sw(new_doc, doc)
            except Exception:
                pass


            # Descrizione su SolidWorks (custom prop)
            try:
                dp = (getattr(self.cfg.solidworks, 'description_prop', 'DESCRIZIONE') or 'DESCRIZIONE').strip()
                if dp and desc:
                    set_custom_properties(doc, {dp: desc})
            except Exception:
                pass
            # Refresh
            self.active_doc_path = str(file_wip)
            self.lbl_doc.configure(text=self._doc_label_text())
            self.code_var.set(code)
            self._refresh_preview()
            self._log_activity("CREATE_CODE", code=code, status="OK", message=f"Creato in WIP ({new_doc.doc_type})")

            messagebox.showinfo("PDM (Macro SolidWorks)", f"Creato codice {code} e salvato in WIP.")

        def _resolve_copy_sources_from_active(self) -> tuple[Path, Path | None, DocType]:
            sw = self._ensure_sw()
            doc = self._active_doc()

            active_path = (self.active_doc_path or "").strip()
            try:
                if doc is not None and hasattr(doc, "GetPathName"):
                    pn = str(doc.GetPathName() or "").strip()
                    if pn:
                        active_path = pn
                        self.active_doc_path = pn
            except Exception:
                pass

            if not active_path:
                raise ValueError("Documento attivo non salvato. Salva il file prima di copiare.")

            p = Path(active_path)
            ext = p.suffix.lower()

            if ext == ".sldprt":
                src_model = p
                src_drw = p.with_suffix(".slddrw")
                doc_type: DocType = "PART"
            elif ext == ".sldasm":
                src_model = p
                src_drw = p.with_suffix(".slddrw")
                doc_type = "ASSY"
            elif ext == ".slddrw":
                src_drw = p
                cand: list[tuple[DocType, Path]] = []
                cand_prt = p.with_suffix(".sldprt")
                cand_asm = p.with_suffix(".sldasm")
                if cand_prt.exists():
                    cand.append(("PART", cand_prt))
                if cand_asm.exists():
                    cand.append(("ASSY", cand_asm))
                if not cand:
                    raise ValueError("Hai aperto un DRW ma non trovo il modello .sldprt/.sldasm omonimo.")
                if len(cand) > 1:
                    raise ValueError("Trovati sia .sldprt che .sldasm con lo stesso nome del DRW. Risolvi il conflitto.")
                doc_type, src_model = cand[0]
            else:
                raise ValueError("Apri un documento PART/ASSY/DRW in SolidWorks.")

            if not src_model.exists():
                raise FileNotFoundError(f"Modello sorgente non trovato: {src_model}")
            if src_drw is not None and (not src_drw.exists()):
                src_drw = None

            # keep SW instance warm; method currently not used but ensures same SW routing
            _ = sw
            return src_model, src_drw, doc_type

        def _get_or_open_sw_doc_by_path(self, sw: Any, file_path: Path) -> tuple[Any, bool]:
            path_s = str(file_path)
            doc = None
            opened_here = False

            try:
                if hasattr(sw, "GetOpenDocumentByName"):
                    doc = sw.GetOpenDocumentByName(path_s)
            except Exception:
                doc = None

            if doc is None:
                try:
                    ad = sw.ActiveDoc
                    if ad is not None and hasattr(ad, "GetPathName"):
                        ap = str(ad.GetPathName() or "")
                        if os.path.normcase(os.path.normpath(ap)) == os.path.normcase(os.path.normpath(path_s)):
                            doc = ad
                except Exception:
                    doc = None

            if doc is None:
                doc = open_doc(sw, path_s, silent=True)
                opened_here = doc is not None

            return doc, opened_here

        def _copy_active_doc_and_save_wip(self):
            try:
                src_model, src_drw, src_doc_type = self._resolve_copy_sources_from_active()
            except Exception as e:
                messagebox.showwarning("PDM (Macro SolidWorks)", f"Copia fallita: {e}")
                return

            src_code = _code_from_path(str(src_model))
            if not src_code:
                messagebox.showwarning("PDM (Macro SolidWorks)", "Impossibile identificare il codice sorgente dal file attivo.")
                return

            src_doc = self.store.get_document(src_code)
            if not src_doc:
                messagebox.showwarning("PDM (Macro SolidWorks)", f"Codice sorgente non presente nel PDM: {src_code}")
                return

            src_state = str(getattr(src_doc, "state", "") or "").strip().upper()
            if src_state not in ("WIP", "REL"):
                messagebox.showwarning("PDM (Macro SolidWorks)", "La copia e consentita solo da documenti in stato WIP o REL.")
                return

            src_doc_type_db = str(getattr(src_doc, "doc_type", "") or "").strip().upper()
            if src_doc_type_db in ("PART", "ASSY"):
                src_doc_type = src_doc_type_db  # trust DB when available

            mmm = (self.mmm_var.get() or "").strip()
            gggg = (self.gggg_var.get() or "").strip()
            vvv = (self.vvv_var.get() or "").strip()
            use_vvv = bool(self.use_vvv_var.get())
            if not use_vvv:
                vvv = ""

            if not mmm or not gggg:
                messagebox.showwarning("PDM (Macro SolidWorks)", "Seleziona MMM e GGGG.")
                return

            try:
                mmm_n = self._normalize_segment("MMM", mmm)
                gggg_n = self._normalize_segment("GGGG", gggg)
                vvv_n = self._normalize_segment("VVV", vvv) if vvv else ""
            except Exception as e:
                messagebox.showwarning("PDM (Macro SolidWorks)", f"Codifica non valida: {e}")
                return

            try:
                seq = self.store.allocate_seq(mmm_n, gggg_n, vvv_n, src_doc_type)
                code = build_code(self.cfg, mmm_n, gggg_n, seq, vvv_n, force_vvv=bool(vvv_n))
            except Exception as e:
                messagebox.showwarning("PDM (Macro SolidWorks)", f"Errore allocazione codice: {e}")
                return

            archive_root = Path(self.cfg.solidworks.archive_root or "").expanduser()
            if not str(archive_root).strip():
                messagebox.showwarning("PDM (Macro SolidWorks)", "Archivio (root) non configurato nel PDM.")
                return

            wip, rel, inrev, _rev = archive_dirs(archive_root, mmm_n, gggg_n)
            file_wip = model_path(wip, code, src_doc_type)
            file_wip_drw = drw_path(wip, code)

            if file_wip.exists():
                messagebox.showwarning("PDM (Macro SolidWorks)", f"File modello destinazione gia presente:\n{file_wip}")
                return
            if src_drw is not None and file_wip_drw.exists():
                messagebox.showwarning("PDM (Macro SolidWorks)", f"File disegno destinazione gia presente:\n{file_wip_drw}")
                return

            desc = str(getattr(src_doc, "description", "") or "").strip().upper()
            self.desc_var.set(desc)

            warnings: list[str] = []
            created_paths: list[Path] = []
            opened_here: list[tuple[Any, str]] = []

            sw = None
            try:
                sw = self._ensure_sw()

                model_doc = None
                try:
                    model_doc, opened = self._get_or_open_sw_doc_by_path(sw, src_model)
                    if model_doc is not None and opened:
                        opened_here.append((model_doc, str(src_model)))
                except Exception:
                    model_doc = None

                model_saved = False
                if model_doc is not None:
                    res_m = save_as_doc(model_doc, str(file_wip))
                    model_saved = bool(getattr(res_m, "ok", False))
                    if not model_saved:
                        msg = str(getattr(res_m, "message", "SaveAs modello fallito."))
                        det = str(getattr(res_m, "details", "") or "").strip()
                        warnings.append("SaveAs modello fallito, uso copia file." + (f" Dettagli: {msg} {det}".strip() if (msg or det) else ""))

                if not model_saved:
                    safe_copy(src_model, file_wip, overwrite=False)
                created_paths.append(file_wip)
                set_readonly(file_wip, readonly=False)

                drw_copied = False
                if src_drw is not None:
                    drw_doc = None
                    try:
                        drw_doc, opened = self._get_or_open_sw_doc_by_path(sw, src_drw)
                        if drw_doc is not None and opened:
                            opened_here.append((drw_doc, str(src_drw)))
                    except Exception:
                        drw_doc = None

                    drw_saved = False
                    if drw_doc is not None:
                        res_d = save_as_doc(drw_doc, str(file_wip_drw))
                        drw_saved = bool(getattr(res_d, "ok", False))
                        if not drw_saved:
                            msg = str(getattr(res_d, "message", "SaveAs disegno fallito."))
                            det = str(getattr(res_d, "details", "") or "").strip()
                            warnings.append("SaveAs disegno fallito, uso copia file." + (f" Dettagli: {msg} {det}".strip() if (msg or det) else ""))

                    if not drw_saved:
                        safe_copy(src_drw, file_wip_drw, overwrite=False)
                    created_paths.append(file_wip_drw)
                    set_readonly(file_wip_drw, readonly=False)
                    drw_copied = True
                else:
                    warnings.append("Disegno sorgente non trovato: copiato solo il modello.")

                new_doc = Document(
                    id=0,
                    code=code,
                    doc_type=src_doc_type,
                    mmm=mmm_n,
                    gggg=gggg_n,
                    seq=seq,
                    vvv=vvv_n,
                    revision=0,
                    state="WIP",
                    obs_prev_state="",
                    description=desc,
                    file_wip_path=str(file_wip),
                    file_rel_path="",
                    file_inrev_path="",
                    file_wip_drw_path=str(file_wip_drw) if drw_copied else "",
                    file_rel_drw_path="",
                    file_inrev_drw_path="",
                    created_at="",
                    updated_at="",
                )
                self.store.add_document(new_doc)

                try:
                    custom_vals = self.store.get_custom_values(src_doc.code) or {}
                except Exception:
                    custom_vals = {}
                for prop_name, prop_value in custom_vals.items():
                    self.store.set_custom_value(code, str(prop_name), str(prop_value))

                _log_line(
                    self.workflow_log_file,
                    "COPY | "
                    f"src={src_doc.code} ({src_state}) -> dst={code} (WIP) | "
                    f"type={src_doc_type} | mmm={mmm_n} gggg={gggg_n} vvv={vvv_n or '-'} | "
                    f"model={file_wip} | drw={'YES' if drw_copied else 'NO'} | custom_props={len(custom_vals)}",
                )

                # allinea proprieta core/descrizione sui file appena creati
                try:
                    mdl_new = open_doc(sw, str(file_wip), silent=True)
                    if mdl_new is not None:
                        self._write_core_props_to_sw(new_doc, mdl_new)
                        dp = (getattr(self.cfg.solidworks, "description_prop", "DESCRIZIONE") or "DESCRIZIONE").strip()
                        if dp:
                            set_custom_properties(mdl_new, {dp: desc})
                        save_existing_doc(mdl_new)
                        close_doc(sw, mdl_new, str(file_wip))
                except Exception:
                    pass

                if drw_copied:
                    try:
                        drw_new = open_doc(sw, str(file_wip_drw), silent=True)
                        if drw_new is not None:
                            self._write_core_props_to_sw(new_doc, drw_new)
                            save_existing_doc(drw_new)
                            close_doc(sw, drw_new, str(file_wip_drw))
                    except Exception:
                        pass

                self.doc_type = src_doc_type
                self.active_doc_path = str(file_wip)
                self.code_var.set(code)
                self.lbl_doc.configure(text=self._doc_label_text())
                self._refresh_preview()
                self._refresh_wf_state()
                self._log_activity("COPY_CODE", code=code, status="OK", message=f"Copia da {src_doc.code} ({src_state})")

                msg = (
                    f"Copia completata.\n\n"
                    f"Sorgente: {src_doc.code} ({src_state})\n"
                    f"Nuovo codice: {code}\n"
                    f"Nuovo stato: WIP | Rev: 00\n"
                    f"Proprieta custom copiate: {len(custom_vals)}"
                )
                messagebox.showinfo("PDM (Macro SolidWorks)", msg)
                if warnings:
                    messagebox.showwarning("PDM (Macro SolidWorks)", "\n".join(warnings))
            except Exception as e:
                self._log_activity("COPY_CODE", code=code if 'code' in locals() else "", status="ERROR", message=str(e))
                for p in reversed(created_paths):
                    try:
                        if p.exists():
                            p.unlink()
                    except Exception:
                        pass
                messagebox.showwarning("PDM (Macro SolidWorks)", f"Copia fallita: {e}")
                return
            finally:
                if sw is not None:
                    for od, p in opened_here:
                        try:
                            close_doc(sw, doc=od, file_path=p)
                        except Exception:
                            pass

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
