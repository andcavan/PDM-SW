from __future__ import annotations

import os
import json
import subprocess
from collections import defaultdict
from datetime import datetime
from pathlib import Path
import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox, filedialog, ttk

import customtkinter as ctk

from pdm_sw.workspace import WorkspaceManager
from pdm_sw.config import ConfigManager, AppConfig, SegmentRule
from pdm_sw.store import Store
from pdm_sw.models import Document
from pdm_sw.codegen import build_code, build_machine_code, build_group_code
from pdm_sw.archive import archive_dirs, archive_dirs_for_machine, archive_dirs_for_group, model_path, drw_path, inrev_tag, safe_copy, set_readonly, release_wip, create_inrev, approve_inrev, cancel_inrev, set_obsolete, restore_obsolete
from pdm_sw.backup import BackupManager
from pdm_sw.sw_integration import test_solidworks_connection
from pdm_sw.macro_publish import publish_macro
from pdm_sw.sw_api import get_solidworks_app, create_model_file, create_drawing_file, open_doc
from pdm_sw.archive_migration import run_archive_layout_migration
from pdm_sw.session_context import resolve_session_context
from pdm_sw.sldreg_manager import import_sldreg_filtered, RestoreOptions as SldregRestoreOptions
from pdm_sw.ui.table import SimpleTable, Table
from pdm_sw.ui.rc_copy_mixin import RCCopyMixin
from pdm_sw.ui.report_mixin import ReportMixin
from pdm_sw.gui.tab_generatore import TabGeneratore
from pdm_sw.gui.tab_codifica import TabCodifica
from pdm_sw.gui.tab_gestione_codifica import TabGestioneCodifica
from pdm_sw.gui.tab_manuale import TabManuale
from pdm_sw.gui.tab_solidworks import TabSolidWorks
from pdm_sw.gui.tab_gerarchia import TabGerarchia
from pdm_sw.gui.tab_monitor import TabMonitor
from pdm_sw.gui.tab_operativo import TabOperativo


APP_DIR = Path(__file__).resolve().parent
LOCAL_SETTINGS_PATH = APP_DIR / "local_settings.json"
APP_REV = "v50.22"
APP_TITLE = f"PDM SolidWorks - Workspace Edition | Rev {APP_REV}"
DOC_LOCK_TTL_SECONDS = 20 * 60
WORKFLOW_WIDTH_RATIO_DEFAULT = 0.40
WORKFLOW_WIDTH_RATIO_MIN = 0.25
WORKFLOW_WIDTH_RATIO_MAX = 0.60

def warn(msg: str) -> None:
    messagebox.showwarning("PDM", msg)


def info(msg: str) -> None:
    messagebox.showinfo("PDM", msg)


def ask(msg: str) -> bool:
    return messagebox.askyesno("PDM", msg)


class PDMApp(RCCopyMixin, ReportMixin, ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.title(APP_TITLE)
        self.geometry("1200x740")

        # Workspace subsystem
        self.shared_root = self._load_shared_data_root()
        self.workspaces_dir = self.shared_root / "WORKSPACES"
        self.ws_mgr = WorkspaceManager(self.workspaces_dir)
        self.ws = self.ws_mgr.ensure_default()
        self.ws_id = self.ws.id

        self.cfg_mgr = ConfigManager(self.ws_mgr.config_path(self.ws_id))
        self.cfg: AppConfig = self.cfg_mgr.load()

        self.store = Store(self.ws_mgr.db_path(self.ws_id))
        self.backup = BackupManager(self.ws_mgr, self.ws_id, self.store, retention_total=self.cfg.backup.retention_total)
        self.session = resolve_session_context()
        self.lock_ttl_seconds = DOC_LOCK_TTL_SECONDS
        self.workflow_width_ratio = self._load_workflow_width_ratio()
        self._operativo_paned = None

                # Chiusura finestra: se _on_close non esiste (merge/override), fallback a destroy
        _handler = getattr(self, "_on_close", None)
        if not callable(_handler):
            _handler = self.destroy
        self.protocol("WM_DELETE_WINDOW", _handler)

        self._build_ui()
        self._refresh_all()
        self._log_activity("APP_START", message=f"Desktop avviato | user_source={self.session.get('source','UNKNOWN')}")

    # ---------------- UI
    def _build_ui(self) -> None:
        # TOP: Workspace bar (modifica GUI richiesta: barra sempre visibile)
        self.ws_bar = ctk.CTkFrame(self)
        self.ws_bar.pack(side="top", fill="x", padx=10, pady=(10, 6))

        self.ws_label = ctk.CTkLabel(self.ws_bar, text="", font=ctk.CTkFont(size=14, weight="bold"))
        self.ws_label.pack(side="left", padx=(12, 8), pady=8)
        self.shared_label = ctk.CTkLabel(self.ws_bar, text="", font=ctk.CTkFont(size=12))
        self.shared_label.pack(side="left", padx=(4, 8), pady=8)

        btn_frame = ctk.CTkFrame(self.ws_bar, fg_color="transparent")
        btn_frame.pack(side="right", padx=8, pady=6)
        self.rev_label = ctk.CTkLabel(
            self.ws_bar,
            text=f"REV {APP_REV}",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.rev_label.pack(side="right", padx=(8, 12), pady=8)
        self.user_label = ctk.CTkLabel(
            self.ws_bar,
            text=f"UTENTE: {self.session.get('display_name', 'unknown')}",
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.user_label.pack(side="right", padx=(8, 8), pady=8)

        self.btn_workspace_tools = ctk.CTkButton(
            btn_frame,
            text="WORKSPACE...",
            command=lambda: self._call_safe("_workspace_tools_dialog"),
            width=160,
        )
        self.btn_workspace_tools.pack(side="left", padx=6)

        # MAIN: tabs
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(side="top", fill="both", expand=True, padx=10, pady=(0, 10))

        self.tab_operativo = self.tabs.add("Operativo")
        self.tab_gerarchia = self.tabs.add("Gerarchia")
        self.tab_monitor = self.tabs.add("Monitor")
        self.tab_setup = self.tabs.add("Setup")

        self.setup_tabs = ctk.CTkTabview(self.tab_setup)
        self.setup_tabs.pack(fill="both", expand=True, padx=6, pady=6)
        self.tab_cod = self.setup_tabs.add("Codifica")
        self.tab_gest_cod = self.setup_tabs.add("Gestione codifica")
        self.tab_gen = self.setup_tabs.add("Generatore codici")
        self.tab_sw = self.setup_tabs.add("SolidWorks")
        self.tab_manuale = self.setup_tabs.add("Manuale")

        # Tab modulari (refactoring completato)
        self.tab_codifica_obj = TabCodifica(self.tab_cod, self, self.cfg, self.store, self.session)
        self.tab_gestione_codifica_obj = TabGestioneCodifica(self.tab_gest_cod, self, self.cfg, self.store, self.session)
        self.tab_generatore_obj = TabGeneratore(self.tab_gen, self, self.cfg, self.store, self.session)
        self.tab_solidworks_obj = TabSolidWorks(self.tab_sw, self, self.cfg, self.store, self.session)
        self.tab_gerarchia_obj = TabGerarchia(self.tab_gerarchia, self, self.cfg, self.store, self.session)
        self.tab_monitor_obj = TabMonitor(self.tab_monitor, self, self.cfg, self.store, self.session)
        self.tab_operativo_obj = TabOperativo(self.tab_operativo, self, self.cfg, self.store, self.session)
        self.tab_manuale_obj = TabManuale(self.tab_manuale, self, self.cfg, self.store, self.session)
        
        # Ponte agli attributi UI workflow per compatibilità con metodi esistenti
        self.rc_table = self.tab_operativo_obj.rc_table
        self.wf_code_var = self.tab_operativo_obj.wf_code_var
        self.wf_state_var = self.tab_operativo_obj.wf_state_var
        self.wf_btn_release = self.tab_operativo_obj.wf_btn_release
        self.wf_btn_create_rev = self.tab_operativo_obj.wf_btn_create_rev
        self.wf_btn_approve = self.tab_operativo_obj.wf_btn_approve
        self.wf_btn_cancel = self.tab_operativo_obj.wf_btn_cancel
        self.wf_btn_obsolete = self.tab_operativo_obj.wf_btn_obsolete
        self.wf_btn_restore_obs = self.tab_operativo_obj.wf_btn_restore_obs
        self.wf_btn_checkout = self.tab_operativo_obj.wf_btn_checkout
        self.wf_btn_checkin = self.tab_operativo_obj.wf_btn_checkin
        self.btn_send_to_wf = self.tab_operativo_obj.btn_send_to_wf
        self.wf_info = self.tab_operativo_obj.wf_info

    def _set_ws_label(self) -> None:
        desc = (self.ws.description or "").strip()
        suffix = f" - {desc}" if desc else ""
        self.ws_label.configure(text=f"WORKSPACE: {self.ws.name}{suffix}")

    def _set_shared_root_label(self) -> None:
        p = str(self.shared_root)
        if len(p) > 65:
            p = "..." + p[-62:]
        self.shared_label.configure(text=f"SHARED: {p}")

    def _norm_segment(self, seg: str, value: str) -> str:
        """Normalizza un segmento (MMM/GGGG/VVV) secondo le regole di 'Gestione codifica'."""
        rule = self.cfg.code.segments.get(seg)
        v = (value or "").strip()
        if not rule:
            return v.upper()
        if not getattr(rule, "enabled", True):
            return v.upper() if rule.case == "UPPER" else v.lower()
        return rule.normalize_value(v)



    def _validate_segment_strict(self, seg: str, value: str, what: str) -> str | None:
        """Valida un segmento secondo regole di Gestione Codifica.

        Regole:
        - UPPER/LOWER viene forzato sempre
        - Se lunghezza e impostata (in eccesso o difetto) -> errore
        - Se charset (ALPHA/NUM/ALNUM) non rispettato -> errore
        - Non applica padding/troncamenti ne rimuove caratteri: o e valido o fallisce.
        """
        rule = self.cfg.code.segments.get(seg)
        v = (value or "").strip()
        if not v:
            return None

        # forza case
        if rule and getattr(rule, "case", "UPPER") == "LOWER":
            v_norm = v.lower()
        else:
            v_norm = v.upper()

        if not rule or not getattr(rule, "enabled", True):
            return v_norm

        # lunghezza esatta
        L = int(getattr(rule, "length", 0) or 0)
        if L > 0 and len(v_norm) != L:
            warn(f"{what} deve essere lungo {L} caratteri. Hai inserito '{v_norm}' ({len(v_norm)}).")
            return None

        charset = getattr(rule, "charset", "ALPHA")
        ok = True
        if charset == "NUM":
            ok = v_norm.isdigit()
        elif charset == "ALPHA":
            ok = v_norm.isalpha()
        else:  # ALNUM
            ok = v_norm.isalnum()

        if not ok:
            warn(f"{what} non rispetta la regola {charset}. Valore inserito: '{v_norm}'.")
            return None

        return v_norm

    def _require_desc_upper(self, desc: str, what: str = "descrizione") -> str | None:
        d = (desc or "").strip().upper()
        if not d:
            warn(f"Inserisci {what}.")
            return None
        return d

    def _ask_large_text_input(
        self,
        title: str,
        prompt: str,
        initial: str = "",
    ) -> str | None:
        result: dict[str, str | None] = {"value": None}

        top = ctk.CTkToplevel(self)
        top.title(title)
        top.geometry("760x230")
        top.grab_set()

        ctk.CTkLabel(
            top,
            text=prompt,
            font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(anchor="w", padx=14, pady=(14, 8))

        value_var = tk.StringVar(value=initial)
        entry = ctk.CTkEntry(
            top,
            textvariable=value_var,
            font=ctk.CTkFont(size=22),
            height=52,
        )
        entry.pack(fill="x", padx=14, pady=(0, 12))
        entry.focus_set()
        try:
            entry.icursor("end")
            entry.select_range(0, "end")
        except Exception:
            pass

        btns = ctk.CTkFrame(top, fg_color="transparent")
        btns.pack(fill="x", padx=14, pady=(0, 12))

        def _cancel():
            result["value"] = None
            top.destroy()

        def _ok():
            result["value"] = value_var.get()
            top.destroy()

        ctk.CTkButton(btns, text="Annulla", width=120, command=_cancel).pack(side="right", padx=6)
        ctk.CTkButton(btns, text="OK", width=120, command=_ok).pack(side="right", padx=6)

        try:
            top.bind("<Return>", lambda _e: _ok())
            top.bind("<Escape>", lambda _e: _cancel())
        except Exception:
            pass

        self.wait_window(top)
        return result["value"]

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

    def _acquire_doc_lock(self, code: str, action: str) -> tuple[bool, dict]:
        ok, lock_status, holder = self.store.acquire_document_lock(
            code=code,
            owner_session=str(self.session.get("session_id", "")),
            owner_user=str(self.session.get("display_name", "")),
            owner_host=str(self.session.get("host", "")),
            ttl_seconds=self.lock_ttl_seconds,
        )
        if ok:
            return True, holder
        who = str(holder.get("owner_user", "") or holder.get("owner_session", "altro utente"))
        host = str(holder.get("owner_host", "") or "")
        lock_msg = f"Documento {code} bloccato da {who}" + (f" su {host}" if host else "") + "."
        self._log_activity(
            action=action,
            code=code,
            status="LOCKED",
            message=lock_status or lock_msg,
            details={"holder": holder},
        )
        warn(lock_msg + "\nRiprova tra poco.")
        return False, holder

    def _release_doc_lock(self, code: str) -> None:
        try:
            self.store.release_document_lock(code=code, owner_session=str(self.session.get("session_id", "")))
        except Exception:
            pass

    def _checkout_identity(self) -> tuple[str, str]:
        user = str(self.session.get("display_name", "") or self.session.get("user_id", "")).strip() or "unknown"
        host = str(self.session.get("host", "")).strip()
        return user, host

    def _is_doc_checked_out_by_me(self, doc: Document | None) -> bool:
        if doc is None:
            return False
        if not bool(getattr(doc, "checked_out", False)):
            return False
        me, _host = self._checkout_identity()
        owner = str(getattr(doc, "checkout_owner_user", "") or "").strip()
        return bool(me) and owner == me

    def _checkout_status_label(self, doc: Document | None) -> str:
        if doc is None:
            return "CHECK-IN"
        if not bool(getattr(doc, "checked_out", False)):
            return "CHECK-IN"
        owner = str(getattr(doc, "checkout_owner_user", "") or "").strip() or "n/d"
        host = str(getattr(doc, "checkout_owner_host", "") or "").strip()
        at = str(getattr(doc, "checkout_at", "") or "").strip().replace("T", " ")
        tail = ""
        if host:
            tail += f" su {host}"
        if at:
            tail += f" @ {at}"
        return f"CHECK-OUT: {owner}{tail}"

    def _checkout_table_value(self, doc: Document | None) -> str:
        if doc is None or not bool(getattr(doc, "checked_out", False)):
            return "IN"
        owner = str(getattr(doc, "checkout_owner_user", "") or "").strip() or "n/d"
        me, _host = self._checkout_identity()
        if me and owner == me:
            return "OUT (ME)"
        return f"OUT ({owner})"

    def _require_wip_checkout(self, doc: Document, action_label: str) -> bool:
        state = str(getattr(doc, "state", "") or "").strip().upper()
        if state != "WIP":
            return True
        if not bool(getattr(doc, "checked_out", False)):
            warn(f"{action_label}: documento WIP non in CHECK-OUT.\nEsegui CHECK-OUT prima di procedere.")
            return False
        if not self._is_doc_checked_out_by_me(doc):
            owner = str(getattr(doc, "checkout_owner_user", "") or "").strip() or "altro utente"
            host = str(getattr(doc, "checkout_owner_host", "") or "").strip()
            who = f"{owner} su {host}" if host else owner
            warn(f"{action_label}: documento in CHECK-OUT da {who}.")
            return False
        return True

    def _require_checkout_for_edit(self, doc: Document, action_label: str) -> bool:
        state = str(getattr(doc, "state", "") or "").strip().upper()
        if state in ("REL", "OBS"):
            warn(f"{action_label}: documento in stato {state}, non modificabile.")
            return False
        if not bool(getattr(doc, "checked_out", False)):
            warn(f"{action_label}: documento non in CHECK-OUT.\nEsegui CHECK-OUT prima di procedere.")
            return False
        if not self._is_doc_checked_out_by_me(doc):
            owner = str(getattr(doc, "checkout_owner_user", "") or "").strip() or "altro utente"
            host = str(getattr(doc, "checkout_owner_host", "") or "").strip()
            who = f"{owner} su {host}" if host else owner
            warn(f"{action_label}: documento in CHECK-OUT da {who}.")
            return False
        return True

    def _get_workflow_or_selected_code(self) -> str:
        try:
            wf_code = (self.wf_code_var.get() or "").strip()
        except Exception:
            wf_code = ""
        if wf_code:
            return wf_code
        code = self._get_table_selected_code()
        if code:
            return code
        try:
            return (self.wf_code_var.get() or "").strip()
        except Exception:
            return ""

    def _checkout_document(self, code: str, show_feedback: bool = True, refresh_ui: bool = True) -> bool:
        code_u = (code or "").strip()
        if not code_u:
            if show_feedback:
                warn("Seleziona un codice.")
            return False
        user, host = self._checkout_identity()
        ok, status, holder = self.store.checkout_document(code_u, owner_user=user, owner_host=host)
        if ok:
            msg = "Documento messo in CHECK-OUT." if status == "CHECKOUT_OK" else "CHECK-OUT gia assegnato a te (refresh)."
            self._log_activity(
                action="CHECKOUT",
                code=code_u,
                status="OK",
                message=msg,
                details={"status_code": status},
            )
            if show_feedback:
                info(msg)
            if refresh_ui:
                self._refresh_all()
            return True

        if status == "CHECKOUT_BY_OTHER":
            owner = str(holder.get("checkout_owner_user", "") or "").strip() or "altro utente"
            host_h = str(holder.get("checkout_owner_host", "") or "").strip()
            msg = f"Documento in CHECK-OUT da {owner}" + (f" su {host_h}" if host_h else "") + "."
            self._log_activity(
                action="CHECKOUT",
                code=code_u,
                status="LOCKED",
                message=msg,
                details={"holder": holder},
            )
            if show_feedback:
                warn(msg)
            return False

        msg = str(status or "CHECK-OUT non riuscito.")
        self._log_activity(
            action="CHECKOUT",
            code=code_u,
            status="WARN",
            message=msg,
            details={"holder": holder},
        )
        if show_feedback:
            warn(msg)
        return False

    def _checkin_document(self, code: str, show_feedback: bool = True, force: bool = False, refresh_ui: bool = True) -> bool:
        code_u = (code or "").strip()
        if not code_u:
            if show_feedback:
                warn("Seleziona un codice.")
            return False
        user, _host = self._checkout_identity()
        ok, status, holder = self.store.checkin_document(code_u, owner_user=user, force=force)
        if ok:
            msg = "Documento messo in CHECK-IN." if status != "CHECKIN_ALREADY" else "Documento gia in CHECK-IN."
            self._log_activity(
                action="CHECKIN",
                code=code_u,
                status="OK",
                message=msg,
                details={"status_code": status},
            )
            if show_feedback:
                info(msg)
            if refresh_ui:
                self._refresh_all()
            return True

        if status == "CHECKOUT_BY_OTHER":
            owner = str(holder.get("checkout_owner_user", "") or "").strip() or "altro utente"
            host_h = str(holder.get("checkout_owner_host", "") or "").strip()
            msg = f"CHECK-IN non consentito: documento in CHECK-OUT da {owner}" + (f" su {host_h}" if host_h else "") + "."
            self._log_activity(
                action="CHECKIN",
                code=code_u,
                status="LOCKED",
                message=msg,
                details={"holder": holder},
            )
            if show_feedback:
                warn(msg)
            return False

        msg = str(status or "CHECK-IN non riuscito.")
        self._log_activity(
            action="CHECKIN",
            code=code_u,
            status="WARN",
            message=msg,
            details={"holder": holder},
        )
        if show_feedback:
            warn(msg)
        return False

    def _checkout_selected_document(self):
        code = self._get_workflow_or_selected_code()
        if not code:
            warn("Seleziona un codice in Consultazione o nel pannello Workflow.")
            return
        self._checkout_document(code, show_feedback=True)

    def _checkin_selected_document(self):
        code = self._get_workflow_or_selected_code()
        if not code:
            warn("Seleziona un codice in Consultazione o nel pannello Workflow.")
            return
        self._checkin_document(code, show_feedback=True)

    def _read_local_settings(self) -> dict:
        try:
            if LOCAL_SETTINGS_PATH.exists():
                d = json.loads(LOCAL_SETTINGS_PATH.read_text(encoding="utf-8"))
                if isinstance(d, dict):
                    return d
        except Exception:
            pass
        return {}

    def _load_shared_data_root(self) -> Path:
        default_root = APP_DIR
        try:
            d = self._read_local_settings()
            p = str((d or {}).get("shared_data_root", "") or "").strip()
            if p:
                return Path(p).expanduser().resolve()
        except Exception:
            pass
        return default_root

    def _clamp_workflow_width_ratio(self, value: float) -> float:
        try:
            v = float(value)
        except Exception:
            v = WORKFLOW_WIDTH_RATIO_DEFAULT
        if v < WORKFLOW_WIDTH_RATIO_MIN:
            v = WORKFLOW_WIDTH_RATIO_MIN
        if v > WORKFLOW_WIDTH_RATIO_MAX:
            v = WORKFLOW_WIDTH_RATIO_MAX
        return v

    def _load_workflow_width_ratio(self) -> float:
        try:
            d = self._read_local_settings()
            v = d.get("workflow_width_ratio", WORKFLOW_WIDTH_RATIO_DEFAULT)
            return self._clamp_workflow_width_ratio(float(v))
        except Exception:
            pass
        return WORKFLOW_WIDTH_RATIO_DEFAULT

    def _save_local_settings(self) -> None:
        try:
            data = {
                "shared_data_root": str(self.shared_root),
                "workflow_width_ratio": float(self._clamp_workflow_width_ratio(self.workflow_width_ratio)),
            }
            LOCAL_SETTINGS_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def _change_shared_root_dialog(self) -> None:
        start = str(self.shared_root if self.shared_root.exists() else APP_DIR)
        picked = filedialog.askdirectory(title="Seleziona cartella dati condivisi", initialdir=start)
        if not picked:
            return
        new_root = Path(picked).expanduser()
        try:
            new_root.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            warn(f"Cartella non accessibile: {e}")
            return
        if new_root.resolve() == self.shared_root.resolve():
            return
        if not ask(f"Impostare cartella condivisa:\n{new_root}\n\nVerranno ricaricati workspace e database. Continuare?"):
            return
        self._switch_shared_root(new_root.resolve())

    def _switch_shared_root(self, new_root: Path) -> None:
        old_root = self.shared_root
        if getattr(self, "monitor_after_id", None):
            try:
                self.after_cancel(self.monitor_after_id)
            except Exception:
                pass
            self.monitor_after_id = None
        try:
            self.store.release_session_locks(str(self.session.get("session_id", "")))
        except Exception:
            pass
        try:
            self.store.close()
        except Exception:
            pass

        self.shared_root = new_root
        self.workspaces_dir = self.shared_root / "WORKSPACES"
        self.ws_mgr = WorkspaceManager(self.workspaces_dir)
        self.ws = self.ws_mgr.ensure_default()
        self.ws_id = self.ws.id

        self.cfg_mgr = ConfigManager(self.ws_mgr.config_path(self.ws_id))
        self.cfg = self.cfg_mgr.load()
        self.store = Store(self.ws_mgr.db_path(self.ws_id))
        self.backup = BackupManager(self.ws_mgr, self.ws_id, self.store, retention_total=self.cfg.backup.retention_total)
        self._rebind_workspace_context_to_tabs(refresh_sw_tab=True)
        self._save_local_settings()
        self._refresh_all()
        self._log_activity("SHARED_ROOT_SWITCH", status="OK", message=f"{old_root} -> {self.shared_root}")
        info(f"Cartella condivisa attiva:\n{self.shared_root}")


    # ---- PDM Custom Properties (Core + Custom) ----
    def _get_custom_prop_defs(self) -> list[dict]:
        try:
            pdm = getattr(self.cfg, "pdm", None)
            defs = list(getattr(pdm, "custom_properties", []) or []) if pdm is not None else []
        except Exception:
            defs = []
        # normalizza: name uppercase
        out = []
        for d in defs:
            if not isinstance(d, dict):
                continue
            name = str(d.get("name", "")).strip().upper()
            if not name:
                continue
            out.append({
                "name": name,
                "type": str(d.get("type", "TEXT") or "TEXT").upper(),
                "required": bool(d.get("required", False)),
                "default": str(d.get("default", "") or ""),
                "options": str(d.get("options", "") or ""),
            })
        return out

    def _get_custom_prop_names(self) -> list[str]:
        return [d["name"] for d in self._get_custom_prop_defs()]

    def _pdm_fields_for_mapping(self) -> list[str]:
        # CORE: proprieta generate dal PDM e inviabili a SolidWorks
        return ["code", "revision", "state", "doc_type", "mmm", "gggg", "vvv"]

    def _sanitize_prop_name(self, name: str) -> str:
        n = (name or "").strip().upper().replace(" ", "_")
        n = "".join(ch for ch in n if (ch.isalnum() or ch == "_"))
        return n

    def _open_manage_pdm_properties(self):
        top = ctk.CTkToplevel(self)
        top.title("Proprieta PDM (Custom)")
        top.geometry("820x520")
        top.grab_set()

        ctk.CTkLabel(top, text="Proprieta PDM (Core + Custom)", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=12, pady=(12, 6))
        ctk.CTkLabel(top, text="Aggiungi o elimina proprieta custom PDM. I nomi sono forzati in MAIUSCOLO.", text_color="#777777").pack(anchor="w", padx=12, pady=(0, 10))

        list_frame = ctk.CTkScrollableFrame(top, height=300)
        list_frame.pack(fill="both", expand=True, padx=12, pady=(0, 10))

        rows = []

        def rebuild():
            for r in list(rows):
                try:
                    r["frame"].destroy()
                except Exception:
                    pass
            rows.clear()

            defs = self._get_custom_prop_defs()
            if not defs:
                ctk.CTkLabel(list_frame, text="(Nessuna proprieta custom definita)").pack(anchor="w", padx=8, pady=8)
                return

            # header
            hdr = ctk.CTkFrame(list_frame, fg_color="transparent")
            hdr.pack(fill="x", pady=(0, 4))
            ctk.CTkLabel(hdr, text="Nome", width=220, anchor="w").pack(side="left", padx=(8, 6))
            ctk.CTkLabel(hdr, text="Tipo", width=120, anchor="w").pack(side="left", padx=6)
            ctk.CTkLabel(hdr, text="Obbligatoria", width=120, anchor="w").pack(side="left", padx=6)
            ctk.CTkLabel(hdr, text="Default", anchor="w").pack(side="left", padx=6)

            for d in defs:
                rf = ctk.CTkFrame(list_frame)
                rf.pack(fill="x", pady=3, padx=2)

                n_var = tk.StringVar(value=d["name"])
                t_var = tk.StringVar(value=d["type"])
                r_var = tk.BooleanVar(value=bool(d["required"]))
                def_var = tk.StringVar(value=d.get("default", ""))

                ctk.CTkEntry(rf, textvariable=n_var, width=220).pack(side="left", padx=(8, 6), pady=6)
                ctk.CTkOptionMenu(rf, variable=t_var, values=["TEXT", "NUM", "DATE", "BOOL", "LIST"], width=120).pack(side="left", padx=6, pady=6)
                ctk.CTkCheckBox(rf, text="", variable=r_var, width=20).pack(side="left", padx=(40, 6), pady=6)
                ctk.CTkEntry(rf, textvariable=def_var).pack(side="left", fill="x", expand=True, padx=6, pady=6)

                def _del(name=n_var.get()):
                    nm = self._sanitize_prop_name(name)
                    if not nm:
                        return
                    if not ask(f"Eliminare proprieta PDM '{nm}'? Verranno cancellati anche i valori salvati nei documenti."):
                        return
                    # rimuovi da config
                    cur = self._get_custom_prop_defs()
                    cur = [x for x in cur if x.get("name","").upper() != nm]
                    self.cfg.pdm.custom_properties = cur
                    self.cfg_mgr.cfg = self.cfg
                    self.cfg_mgr.save()
                    # elimina valori DB
                    try:
                        self.store.delete_custom_property_values(nm)
                    except Exception:
                        pass
                    rebuild()
                    # aggiorna UI collegate
                    try:
                        self._update_sw_mapping_field_values()
                    except Exception:
                        pass
                    try:
                        self._refresh_custom_props_inputs()
                    except Exception:
                        pass

                ctk.CTkButton(rf, text="X", width=34, command=_del).pack(side="left", padx=6, pady=6)

                rows.append({"frame": rf, "name": n_var, "type": t_var, "req": r_var, "default": def_var})

        # Add area
        add_frame = ctk.CTkFrame(top)
        add_frame.pack(fill="x", padx=12, pady=(0, 10))
        ctk.CTkLabel(add_frame, text="Nuova proprieta").grid(row=0, column=0, padx=8, pady=8, sticky="w")
        new_name = tk.StringVar(value="")
        new_type = tk.StringVar(value="TEXT")
        new_req = tk.BooleanVar(value=False)
        new_def = tk.StringVar(value="")
        ctk.CTkEntry(add_frame, textvariable=new_name, width=220).grid(row=0, column=1, padx=6, pady=8, sticky="w")
        ctk.CTkOptionMenu(add_frame, variable=new_type, values=["TEXT", "NUM", "DATE", "BOOL", "LIST"], width=120).grid(row=0, column=2, padx=6, pady=8, sticky="w")
        ctk.CTkCheckBox(add_frame, text="Obbligatoria", variable=new_req).grid(row=0, column=3, padx=10, pady=8, sticky="w")
        ctk.CTkEntry(add_frame, textvariable=new_def).grid(row=0, column=4, padx=6, pady=8, sticky="ew")
        add_frame.grid_columnconfigure(4, weight=1)

        def do_add():
            nm = self._sanitize_prop_name(new_name.get())
            if not nm:
                warn("Inserisci un nome proprieta valido.")
                return
            cur_names = set(self._get_custom_prop_names())
            if nm in cur_names:
                warn(f"La proprieta '{nm}' esiste gia.")
                return
            entry = {"name": nm, "type": (new_type.get() or "TEXT").upper(), "required": bool(new_req.get()), "default": str(new_def.get() or ""), "options": ""}
            cur = self._get_custom_prop_defs()
            cur.append(entry)
            self.cfg.pdm.custom_properties = cur
            self.cfg_mgr.cfg = self.cfg
            self.cfg_mgr.save()
            new_name.set(""); new_type.set("TEXT"); new_req.set(False); new_def.set("")
            rebuild()
            try:
                self._update_sw_mapping_field_values()
            except Exception:
                pass
            try:
                self._refresh_custom_props_inputs()
            except Exception:
                pass

        ctk.CTkButton(add_frame, text="Aggiungi", width=120, command=do_add).grid(row=0, column=5, padx=8, pady=8)

        bottom = ctk.CTkFrame(top, fg_color="transparent")
        bottom.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(bottom, text="Chiudi", width=120, command=top.destroy).pack(side="right")

        rebuild()

    def _update_sw_mapping_field_values(self):
        """Aggiorna i valori disponibili nei menu 'Campo PDM' della mappatura SolidWorks."""
        if not hasattr(self, "sw_map_rows"):
            return
        values = self._pdm_fields_for_mapping()
        for r in (self.sw_map_rows or []):
            opt = r.get("opt")
            var = r.get("field_var")
            if opt is None or var is None:
                continue
            try:
                opt.configure(values=values)
                cur = var.get()
                if cur not in values:
                    var.set("code")
            except Exception:
                pass

    def _refresh_custom_props_inputs(self):
        """Ricrea gli input delle proprieta custom nella tab Codifica."""
        if not hasattr(self, "custom_props_frame") or self.custom_props_frame is None:
            return
        frame = self.custom_props_frame
        for ch in frame.winfo_children():
            try:
                ch.destroy()
            except Exception:
                pass
        self.custom_prop_vars = {}
        defs = self._get_custom_prop_defs()
        if not defs:
            ctk.CTkLabel(frame, text="(Nessuna proprieta custom definita)", text_color="#777777").pack(anchor="w", padx=8, pady=6)
            return

        # grid with 2 columns
        grid = ctk.CTkFrame(frame, fg_color="transparent")
        grid.pack(fill="x", padx=6, pady=6)

        row = 0
        col = 0
        for d in defs:
            name = d["name"]
            req = bool(d.get("required", False))
            default = str(d.get("default", "") or "")
            v = tk.StringVar(value=default)
            self.custom_prop_vars[name] = v

            cell = ctk.CTkFrame(grid, fg_color="transparent")
            cell.grid(row=row, column=col, padx=8, pady=6, sticky="ew")
            lbl = f"{name}{' *' if req else ''}"
            ctk.CTkLabel(cell, text=lbl, anchor="w").pack(anchor="w")
            ctk.CTkEntry(cell, textvariable=v, width=240).pack(fill="x")

            col += 1
            if col >= 2:
                col = 0
                row += 1

        grid.grid_columnconfigure(0, weight=1)
        grid.grid_columnconfigure(1, weight=1)

    def _save_custom_props_for_doc(self, code: str):
        defs = self._get_custom_prop_defs()
        if not defs or not hasattr(self, "custom_prop_vars"):
            return
        for d in defs:
            name = d["name"]
            req = bool(d.get("required", False))
            v = ""
            try:
                v = (self.custom_prop_vars.get(name).get() if self.custom_prop_vars.get(name) else "")
            except Exception:
                v = ""
            v = (v or "").strip()
            if req and not v:
                raise ValueError(f"La proprieta '{name}' e obbligatoria.")
            # salva anche vuoto? preferisco salvare solo se valorizzato o se required
            if v or req or str(d.get("default","") or ""):
                self.store.set_custom_value(code, name, v)

    # ---------------- Tab: Gestione codifica
    # ===============================================================================
    # METODO DEPRECATO - Estratto in pdm_sw.gui.tab_gestione_codifica.TabGestioneCodifica
    # Mantenuto temporaneamente per riferimento durante il refactoring
    # ===============================================================================
    # def _ui_gestione_codifica(self):
    #     frame = ctk.CTkScrollableFrame(self.tab_gest_cod)
    #     frame.pack(fill="both", expand=True, padx=10, pady=10)
    #
    #     ctk.CTkLabel(frame, text="Configurazione codifica [MMM]_[GGGG]-[VVV]-[0000]", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", pady=(0, 10))
    #
    #     row1 = ctk.CTkFrame(frame)
    #     row1.pack(fill="x", pady=6)
    #
    #     self.sep1_var = tk.StringVar(value=self.cfg.code.sep1)
    #     self.sep2_var = tk.StringVar(value=self.cfg.code.sep2)
    #     self.sep3_var = tk.StringVar(value=self.cfg.code.sep3)
    #     self.include_vvv_var = tk.BooleanVar(value=self.cfg.code.include_vvv_by_default)
    #
    #     ctk.CTkLabel(row1, text="Separatore MMM/GGGG").pack(side="left", padx=(8, 6))
    #     ctk.CTkEntry(row1, width=60, textvariable=self.sep1_var).pack(side="left", padx=6)
    #     ctk.CTkLabel(row1, text="Separatore GGGG/0000").pack(side="left", padx=(18, 6))
    #     ctk.CTkEntry(row1, width=60, textvariable=self.sep2_var).pack(side="left", padx=6)
    #     ctk.CTkLabel(row1, text="Separatore 0000/VVV").pack(side="left", padx=(18, 6))
    #     ctk.CTkEntry(row1, width=60, textvariable=self.sep3_var).pack(side="left", padx=6)
    #
    #     row2 = ctk.CTkFrame(frame)
    #     row2.pack(fill="x", pady=6)
    #     ctk.CTkCheckBox(row2, text="Includi VVV di default", variable=self.include_vvv_var).pack(side="left", padx=8)
    #
    #     row3 = ctk.CTkFrame(frame)
    #     row3.pack(fill="x", pady=6)
    #     ctk.CTkLabel(row3, text="Preset VVV (separati da virgola)").pack(side="left", padx=(8, 6))
    #     self.vvv_var = tk.StringVar(value=",".join(self.cfg.code.vvv_presets))
    #     ctk.CTkEntry(row3, textvariable=self.vvv_var).pack(side="left", fill="x", expand=True, padx=(6, 8))
    #
    #
    #     # Segment rules
    #     ctk.CTkLabel(frame, text="Regole segmenti", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", pady=(16, 6))
    #
    #     seg_frame = ctk.CTkFrame(frame)
    #     seg_frame.pack(fill="x", pady=6)
    #
    #     # headers
    #     hdr = ctk.CTkFrame(seg_frame, fg_color="transparent")
    #     hdr.pack(fill="x", padx=8, pady=(8, 2))
    #     ctk.CTkLabel(hdr, text="Segmento", width=90).pack(side="left")
    #     ctk.CTkLabel(hdr, text="Attivo", width=70).pack(side="left")
    #     ctk.CTkLabel(hdr, text="Lunghezza", width=90).pack(side="left")
    #     ctk.CTkLabel(hdr, text="Charset", width=120).pack(side="left")
    #     ctk.CTkLabel(hdr, text="Case", width=120).pack(side="left")
    #
    #     self.seg_rule_vars = {}
    #     for token in ["MMM", "GGGG", "0000", "VVV"]:
    #         rule = self.cfg.code.segments.get(token)
    #         row = ctk.CTkFrame(seg_frame, fg_color="transparent")
    #         row.pack(fill="x", padx=8, pady=2)
    #
    #         enabled_var = tk.BooleanVar(value=(rule.enabled if rule else True))
    #         length_var = tk.StringVar(value=str(rule.length if rule else (4 if token == "0000" else len(token))))
    #         charset_var = tk.StringVar(value=(rule.charset if rule else ("NUM" if token == "0000" else "ALPHA")))
    #         case_var = tk.StringVar(value=(rule.case if rule else "UPPER"))
    #
    #         self.seg_rule_vars[token] = {
    #             "enabled": enabled_var,
    #             "length": length_var,
    #             "charset": charset_var,
    #             "case": case_var,
    #         }
    #
    #         ctk.CTkLabel(row, text=token, width=90).pack(side="left")
    #         ctk.CTkCheckBox(row, text="", variable=enabled_var, width=70, command=self._refresh_code_config_preview).pack(side="left")
    #         ctk.CTkEntry(row, width=90, textvariable=length_var).pack(side="left", padx=(0, 10))
    #         ctk.CTkOptionMenu(row, width=120, values=["NUM", "ALPHA", "ALNUM"], variable=charset_var, command=lambda _=None: self._refresh_code_config_preview()).pack(side="left", padx=(0, 10))
    #         ctk.CTkOptionMenu(row, width=120, values=["UPPER", "LOWER"], variable=case_var, command=lambda _=None: self._refresh_code_config_preview()).pack(side="left")
    #
    #     self.code_cfg_preview = ctk.CTkLabel(frame, text="")
    #     self.code_cfg_preview.pack(anchor="w", pady=(10, 0))
    #     self._refresh_code_config_preview()
    #
    #     btns = ctk.CTkFrame(frame, fg_color="transparent")
    #     btns.pack(fill="x", pady=12)
    #     ctk.CTkButton(btns, text="Salva configurazione", command=self._save_code_config).pack(side="left", padx=8)
    #     self.code_preview = ctk.CTkLabel(btns, text="")
    #     self.code_preview.pack(side="left", padx=14)

    # def _save_code_config(self):
    #     self.cfg.code.sep1 = self.sep1_var.get()
    #     self.cfg.code.sep2 = self.sep2_var.get()
    #     self.cfg.code.sep3 = self.cfg.code.sep3 = self.sep3_var.get()
    #     self.cfg.code.include_vvv_by_default = bool(self.include_vvv_var.get())
    #     self.cfg.code.vvv_presets = [x.strip() for x in self.vvv_var.get().split(",") if x.strip()]
    #
    #     # Segment rules
    #     if hasattr(self, "seg_rule_vars") and self.seg_rule_vars:
    #         for token, vars_ in self.seg_rule_vars.items():
    #             try:
    #                 length = int(str(vars_["length"].get()).strip() or "0")
    #             except Exception:
    #                 length = 0
    #             length = max(1, min(12, length)) if token != "0000" else max(1, min(6, length))
    #             rule = self.cfg.code.segments.get(token) or SegmentRule()
    #             rule.enabled = bool(vars_["enabled"].get())
    #             rule.length = length
    #             rule.charset = str(vars_["charset"].get())
    #             rule.case = str(vars_["case"].get())
    #             self.cfg.code.segments[token] = rule
    #     self._refresh_code_config_preview()
    #     self.cfg_mgr.cfg = self.cfg
    #     self.cfg_mgr.save()


    def _default_sw_property_map(self) -> dict:
        # Default nomi proprieta (italiano) - modificabili nella tab SolidWorks
        return {
            "code": "CODICE",
                        "revision": "REVISIONE",
            "state": "STATO",
            "doc_type": "TIPO_DOC",
            "mmm": "MACCHINA",
            "gggg": "GRUPPO",
            "vvv": "VARIANTE",
        }

    def _build_sw_props_for_doc(self, doc) -> dict:
        # Usa lista mapping se presente; fallback su dict legacy
        try:
            mappings = list(getattr(self.cfg.solidworks, "property_mappings", []) or [])
        except Exception:
            mappings = []
        mp = dict(getattr(self.cfg.solidworks, "property_map", {}) or {})
        props = {}

        def get_pdm_value(field: str):
            if str(field).startswith("custom:"):
                # Le proprieta custom sono gestite da SolidWorks (SW->PDM). Non scrivere verso SW.
                return None
            if field == "code":
                return doc.code
            if field == "description":
                return (doc.description or "").upper()
            if field == "revision":
                try:
                    return f"{int(doc.revision):02d}"
                except Exception:
                    return str(doc.revision)
            if field == "state":
                return doc.state
            if field == "doc_type":
                return doc.doc_type
            if field == "mmm":
                return doc.mmm
            if field == "gggg":
                return doc.gggg
            if field == "vvv":
                return (doc.vvv or "")
            return ""

        iterable = []
        if mappings:
            iterable = [(str(it.get('pdm_field','') or it.get('field','') or ''), str(it.get('sw_prop','') or it.get('sw','') or '')) for it in mappings]
        else:
            iterable = list(mp.items())

        for field, sw_prop in iterable:
            sw_prop = (sw_prop or "").strip()
            if not sw_prop:
                continue
            v = get_pdm_value(field)
            if v is None:
                continue
            props[sw_prop] = str(v)
        return props
        self._refresh_vvv_menu()
        info("Configurazione codifica salvata.")


    # def _refresh_code_config_preview(self):
    #     # Anteprima basata sui valori inseriti (non salva automaticamente)
    #     try:
    #         sep1 = self.sep1_var.get()
    #         sep2 = self.sep2_var.get()
    #         sep3 = self.sep3_var.get()
    #         include_vvv = bool(self.include_vvv_var.get())
    #         vvv_sample = (self.cfg.code.vvv_presets[0] if self.cfg.code.vvv_presets else "V01")
    #
    #         # build temporary segment rules
    #         segs = {}
    #         if hasattr(self, "seg_rule_vars"):
    #             for token, vars_ in self.seg_rule_vars.items():
    #                 try:
    #                     length = int(str(vars_["length"].get()).strip() or "0")
    #                 except Exception:
    #                     length = 0
    #                 length = max(1, min(12, length)) if token != "0000" else max(1, min(6, length))
    #                 segs[token] = SegmentRule(
    #                     enabled=bool(vars_["enabled"].get()),
    #                     length=length,
    #                     charset=str(vars_["charset"].get()),
    #                     case=str(vars_["case"].get()),
    #                 )
    #         else:
    #             segs = self.cfg.code.segments
    #
    #         mmm_v = segs["MMM"].normalize_value("MAC") if segs.get("MMM") and segs["MMM"].enabled else ""
    #         gggg_v = segs["GGGG"].normalize_value("GRUP") if segs.get("GGGG") and segs["GGGG"].enabled else ""
    #         seq_len = segs["0000"].length if segs.get("0000") else 4
    #         seq_v = str(1).zfill(seq_len) if (segs.get("0000") and segs["0000"].enabled) else ""
    #         code = mmm_v + sep1 + gggg_v + sep2
    #         if include_vvv:
    #             vvv_v = segs["VVV"].normalize_value(vvv_sample) if segs.get("VVV") and segs["VVV"].enabled else ""
    #             if vvv_v:
    #                 code += vvv_v + sep3 + seq_v
    #             else:
    #                 code += seq_v
    #         else:
    #             code += seq_v
    #
    #         if hasattr(self, "code_cfg_preview"):
    #             self.code_cfg_preview.configure(text=f"Anteprima: {code}")
    #         if hasattr(self, "code_preview"):
    #             self.code_preview.configure(text=f"Anteprima: {code}")
    #     except Exception:
    #         pass


    # ---------------- Tab: Generatore codici (macchine/gruppi)
    # ===============================================================================
    # METODO DEPRECATO - Estratto in pdm_sw.gui.tab_generatore.TabGeneratore
    # Mantenuto temporaneamente per riferimento durante il refactoring
    # ===============================================================================
    # ===============================================================================
    # METODO DEPRECATO - Estratto in pdm_sw.gui.tab_generatore.TabGeneratore
    # Tutto il codice UI originale è stato commentato per evitare duplicazioni
    # ===============================================================================
    # def _ui_generatore(self):
    #     outer = ctk.CTkFrame(self.tab_gen)
    #     outer.pack(fill="both", expand=True, padx=10, pady=10)
    #     ... (codice rimosso, ora gestito da TabGeneratore)

    def _on_machine_list_selected(self):
        mmm = self._selected_mmm()
        if hasattr(self, "group_mmm_var"):
            self.group_mmm_var.set(mmm)
        self._refresh_groups()

    def _selected_mmm(self) -> str:
        sel = self.machine_list.curselection()
        if sel:
            val = self.machine_list.get(sel[0])
            return val.split(" ")[0].strip()
        # fallback: usa il menu macchina quando la listbox non ha selezione attiva
        if hasattr(self, "group_mmm_var"):
            return (self.group_mmm_var.get() or "").strip().upper()
        return ""

    def _selected_gggg(self) -> str:
        sel = self.group_list.curselection()
        if sel:
            val = self.group_list.get(sel[0])
            return val.split(" ")[0].strip()
        # fallback: riga attiva (utile quando il focus passa al pulsante Modifica)
        try:
            idx = int(self.group_list.index(tk.ACTIVE))
            if 0 <= idx < int(self.group_list.size()):
                val = self.group_list.get(idx)
                return val.split(" ")[0].strip()
        except Exception:
            pass
        return ""

    def _refresh_machines(self):
        """Delega al TabGeneratore se disponibile."""
        if hasattr(self, 'tab_generatore_obj') and self.tab_generatore_obj:
            self.tab_generatore_obj.refresh_machines()
        # Altrimenti skip (tab non ancora inizializzato)


    def _refresh_groups(self):
        """Delega al TabGeneratore se disponibile."""
        if hasattr(self, 'tab_generatore_obj') and self.tab_generatore_obj:
            self.tab_generatore_obj.refresh_groups()
        # Altrimenti skip (tab non ancora inizializzato)


    def _on_group_machine_selected(self):
        # chiamato dal menu a tendina nella sezione gruppi
        mmm = (self.group_mmm_var.get() or "").strip().upper() if hasattr(self, "group_mmm_var") else ""
        if not mmm:
            self.group_list.delete(0, tk.END)
            return
        # sincronizza selezione listbox macchine se possibile
        try:
            for i in range(self.machine_list.size()):
                if str(self.machine_list.get(i)).split(" ")[0].strip() == mmm:
                    self.machine_list.selection_clear(0, tk.END)
                    self.machine_list.selection_set(i)
                    self.machine_list.see(i)
                    break
        except Exception:
            pass
        self._refresh_groups()
        self._refresh_group_menu()

    def _add_machine(self):
        mmm = self._validate_segment_strict("MMM", self.mmm_new.get(), "MMM")
        if not mmm:
            warn("Inserisci MMM.")
            return
        name = self._require_desc_upper(self.mmm_name_new.get(), what="descrizione macchina")
        if name is None:
            return
        self.store.add_machine(mmm, name)
        self.mmm_new.set(""); self.mmm_name_new.set("")
        self._refresh_machines()
        self._refresh_machine_menus()
        self._refresh_hierarchy_tree()

    def _edit_machine_desc(self):
        mmm = self._selected_mmm()
        if not mmm:
            warn("Seleziona una macchina.")
            return

        current_name = ""
        for m, name in self.store.list_machines():
            if str(m) == str(mmm):
                current_name = str(name or "")
                break

        new_name_raw = self._ask_large_text_input(
            title="Modifica descrizione macchina",
            prompt=f"Nuova descrizione per macchina {mmm}:",
            initial=current_name,
        )
        if new_name_raw is None:
            return

        new_name = self._require_desc_upper(new_name_raw, what="descrizione macchina")
        if new_name is None:
            return

        self.store.add_machine(mmm, new_name)
        self._refresh_machines()
        if hasattr(self, "group_mmm_var"):
            self.group_mmm_var.set(mmm)
        self._on_group_machine_selected()
        self._refresh_machine_menus()
        self._refresh_hierarchy_tree()

    def _del_machine(self):
        mmm = self._selected_mmm()
        if not mmm:
            return
        if not ask(f"Eliminare macchina {mmm}?"):
            return
        self.store.delete_machine(mmm)
        self._refresh_machines()
        self.group_list.delete(0, tk.END)
        self._refresh_machine_menus()
        self._refresh_hierarchy_tree()

    def _add_group(self):
        mmm = (self.group_mmm_var.get().strip() if hasattr(self, 'group_mmm_var') else self._selected_mmm())
        mmm = self._validate_segment_strict("MMM", mmm, "MMM")
        if not mmm:
            warn("Seleziona una macchina.")
            return

        gggg = self._validate_segment_strict("GGGG", self.gggg_new.get(), "GGGG")
        if not gggg:
            warn("Inserisci GGGG.")
            return

        name = self._require_desc_upper(self.gggg_name_new.get(), what="descrizione gruppo")
        if name is None:
            return

        self.store.add_group(mmm, gggg, name)
        self.gggg_new.set(""); self.gggg_name_new.set("")
        self._refresh_groups()
        self._refresh_group_menu()
        self._refresh_hierarchy_tree()

    def _edit_group_desc(self):
        mmm = (self.group_mmm_var.get().strip().upper() if hasattr(self, 'group_mmm_var') else self._selected_mmm())
        gggg = self._selected_gggg()
        if not mmm or not gggg:
            warn("Seleziona un gruppo.")
            return

        current_name = ""
        for g, name in self.store.list_groups(mmm):
            if str(g) == str(gggg):
                current_name = str(name or "")
                break

        new_name_raw = self._ask_large_text_input(
            title="Modifica descrizione gruppo",
            prompt=f"Nuova descrizione per gruppo {mmm}/{gggg}:",
            initial=current_name,
        )
        if new_name_raw is None:
            return

        new_name = self._require_desc_upper(new_name_raw, what="descrizione gruppo")
        if new_name is None:
            return

        self.store.add_group(mmm, gggg, new_name)
        self._refresh_groups()
        self._refresh_group_menu()
        self._refresh_hierarchy_tree()

    def _del_group(self):
        mmm = (self.group_mmm_var.get().strip().upper() if hasattr(self, 'group_mmm_var') else self._selected_mmm())
        sel = self.group_list.curselection()
        if not mmm or not sel:
            return
        gggg = self.group_list.get(sel[0]).split(" ")[0].strip()
        if not ask(f"Eliminare gruppo {mmm}/{gggg}?"):
            return
        self.store.delete_group(mmm, gggg)
        self._refresh_groups()
        self._refresh_group_menu()
        self._refresh_hierarchy_tree()

    # ===============================================================================
    # METODO DEPRECATO - Estratto in pdm_sw.gui.tab_solidworks.TabSolidWorks
    # ===============================================================================
    # def _ui_solidworks(self):
    #     frame = ctk.CTkScrollableFrame(self.tab_sw)
    #     frame.pack(fill="both", expand=True, padx=10, pady=10)

    #     ctk.CTkLabel(frame, text="Impostazioni SolidWorks", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", pady=(0, 10))
    # 
    #     def row(label: str, var: tk.StringVar, browse: bool = True, is_file: bool = False):
    #         r = ctk.CTkFrame(frame)
    #         r.pack(fill="x", pady=6)
    #         ctk.CTkLabel(r, text=label, width=170, anchor="w").pack(side="left", padx=(8, 6))
    #         ctk.CTkEntry(r, textvariable=var).pack(side="left", fill="x", expand=True, padx=6)
    #         if browse:
    #             def _pick():
    #                 if is_file:
    #                     p = filedialog.askopenfilename()
    #                 else:
    #                     p = filedialog.askdirectory()
    #                 if p:
    #                     var.set(p)
    #             ctk.CTkButton(r, text="Sfoglia", width=90, command=_pick).pack(side="left", padx=6)
    # 
    #     self.archive_root_var = tk.StringVar(value=self.cfg.solidworks.archive_root)
    #     self.tpl_part_var = tk.StringVar(value=self.cfg.solidworks.template_part)
    #     self.tpl_assy_var = tk.StringVar(value=self.cfg.solidworks.template_assembly)
    #     self.tpl_drw_var = tk.StringVar(value=self.cfg.solidworks.template_drawing)
    # 
    #     row("Archivio (root)", self.archive_root_var, browse=True, is_file=False)
    #     row("Template PART", self.tpl_part_var, browse=True, is_file=True)
    #     row("Template ASSY", self.tpl_assy_var, browse=True, is_file=True)
    #     row("Template DRW", self.tpl_drw_var, browse=True, is_file=True)
    #     ... (resto del codice UI estratto in TabSolidWorks)

    # def _save_sw_config(self):
    #     ... (estratto in TabSolidWorks._save_sw_config)

    # def _test_sw(self):
    #     ... (estratto in TabSolidWorks._test_sw)

    # def _publish_sw_macro(self) -> None:
    #    ... (estratto in TabSolidWorks._publish_sw_macro)

    # ---------------- Tab: Codifica (creazione documenti)
    # ===============================================================================
    # METODO DEPRECATO - Estratto in pdm_sw.gui.tab_codifica.TabCodifica
    # Mantenuto temporaneamente per riferimento durante il refactoring
    # ===============================================================================
    # ===============================================================================
    # METODO DEPRECATO - Estratto in pdm_sw.gui.tab_codifica.TabCodifica
    # Tutto il codice UI originale è stato commentato per evitare duplicazioni
    # ===============================================================================
    # def _ui_codifica(self):
    #     frame = ctk.CTkFrame(self.tab_cod)
    #     frame.pack(fill="both", expand=True, padx=10, pady=10)
    #     ... (codice rimosso, ora gestito da TabCodifica)


    # def _publish_sw_macro(self) -> None:
    #     """Pubblica (genera) la macro di bootstrap SolidWorks + payload per la workspace corrente."""
    #     try:
    #         bas_path, payload_dir = publish_macro(APP_DIR, self.ws_id)
    #         info(
    #             "Macro SolidWorks pubblicata.\n\n"
    #             f"Bootstrap (.bas): {bas_path}\n"
    #             f"Payload: {payload_dir}\n\n"
    #             "Apri il file di istruzioni nella cartella SW_MACROS (INSTALL_MACRO_<workspace>.txt)."
    #         )
    #     except Exception as e:
    #         warn(f"Pubblicazione macro fallita: {e}")

    def _refresh_machine_menus(self):
        """Delega al TabCodifica se disponibile."""
        if hasattr(self, 'tab_codifica_obj') and self.tab_codifica_obj:
            self.tab_codifica_obj.refresh_machine_menus()
        # Altrimenti skip (tab non ancora inizializzato)

    def _refresh_group_menu(self):
        """Delega al TabCodifica se disponibile."""
        if hasattr(self, 'tab_codifica_obj') and self.tab_codifica_obj:
            self.tab_codifica_obj._refresh_group_menu()
        # Altrimenti skip (tab non ancora inizializzato)

    def _refresh_vvv_menu(self):
        """Delega al TabCodifica se disponibile."""
        if hasattr(self, 'tab_codifica_obj') and self.tab_codifica_obj:
            self.tab_codifica_obj.refresh_vvv_menu()
        # Altrimenti skip (tab non ancora inizializzato)

    # ===============================================================================
    # METODI DEPRECATI - Estratti in TabCodifica (accedono a widget del tab)
    # ===============================================================================
    # def _on_doc_type_change(self):
    # def _refresh_preview(self):
    # def _show_next_code(self):
    # def _browse_link_file(self):
    # def _clear_link_file(self):
    # def _generate_document(self):
    #     ... (codice rimosso, ora gestito da TabCodifica)

    # Nota: _import_linked_files_to_wip() è ancora usato dal TabCodifica come helper

    # ===============================================================================
    # METODI RIMOSSI - Estratti in TabCodifica (accedevano a widget del tab)
    # Tutti i seguenti metodi sono ora gestiti da TabCodifica:
    # - _on_doc_type_change()
    # - _refresh_preview()
    # - _show_next_code()
    # - _browse_link_file()
    # - _clear_link_file()
    # - _generate_document()
    # ===============================================================================

    def _import_linked_files_to_wip(self, doc: Document, src_path: str, auto_drw: bool = True) -> None:
        # Copia un file esistente (e relativo DRW se presente) nella cartella WIP dell'archivio
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
            raise ValueError("Il file selezionato e una PARTE (.sldprt) ma il tipo scelto e ASSY.")
        if src_model.suffix.lower() == '.sldasm' and doc.doc_type != 'ASSY':
            raise ValueError("Il file selezionato e un ASSIEME (.sldasm) ma il tipo scelto e PART.")

        wip, rel, inrev, rev = archive_dirs(self.cfg.solidworks.archive_root, doc.mmm, doc.gggg)
        dst_model = model_path(wip, doc.code, doc.doc_type)
        if dst_model.exists():
            raise ValueError("Esiste gia un file modello in archivio con questo codice (WIP).")
        safe_copy(src_model, dst_model)
        set_readonly(dst_model, readonly=False)
        self.store.update_document(doc.code, file_wip_path=str(dst_model))

        # DRW: stesso nome del modello, stessa cartella (se auto_drw è True)
        if auto_drw:
            src_drw = src_model.with_suffix('.slddrw')
            if src_drw.exists():
                dst_drw = drw_path(wip, doc.code)
                if not dst_drw.exists():
                    safe_copy(src_drw, dst_drw)
                    set_readonly(dst_drw, readonly=False)
                self.store.update_document(doc.code, file_wip_drw_path=str(dst_drw))

    # Metodo _generate_document() rimosso - ora gestito da TabCodifica

    def _create_files_for_code(self, code: str, create_drw: bool | None = None, only_missing: bool = False, require_checkout: bool = True):
        doc = self.store.get_document(code)
        if not doc:
            return
        if str(doc.state or "").strip().upper() != "WIP":
            warn("Creazione file consentita solo per documenti in stato WIP.")
            return
        if require_checkout and not self._require_checkout_for_edit(doc, "Creazione file"):
            return
        if not self.cfg.solidworks.archive_root:
            warn("Archivio non impostato (tab SolidWorks).")
            return

        tpl_model = self.cfg.solidworks.template_part if doc.doc_type == "PART" else self.cfg.solidworks.template_assembly
        wip, rel, inrev, rev = archive_dirs(self.cfg.solidworks.archive_root, doc.mmm, doc.gggg)
        out_model = model_path(wip, doc.code, doc.doc_type)
        out_drw = drw_path(wip, doc.code)

        if create_drw is None:
            try:
                create_drw = bool(self.create_drw_var.get())
            except Exception:
                create_drw = True

        model_exists = out_model.is_file()
        drw_exists = out_drw.is_file()
        need_model = (not only_missing) or (not model_exists)
        need_drw = bool(create_drw) and ((not only_missing) or (not drw_exists))

        if not need_model and not need_drw:
            self.store.update_document(doc.code, file_wip_path=str(out_model), file_wip_drw_path=str(out_drw))
            info("Nessun file mancante: MODEL e DRW gia presenti in WIP.")
            self._refresh_all()
            return

        if need_model and not tpl_model:
            warn("Template modello non impostato (tab SolidWorks).")
            return
        if need_drw and not self.cfg.solidworks.template_drawing:
            if need_model:
                warn("Template DRW non impostato: salto creazione disegno.")
                need_drw = False
            else:
                warn("Template DRW non impostato (tab SolidWorks).")
                return

        self._apply_sldreg_before_sw_launch(code=doc.code, open_source="CREATE_FILES", kind="MODEL")
        sw, res = get_solidworks_app(visible=False, timeout_s=30.0)
        if not res.ok or sw is None:
            warn(res.message + ("\n\n" + res.details if res.details else ""))
            return

        props = self._build_sw_props_for_doc(doc)
        # seed descrizione in SolidWorks (poi sara gestita da SolidWorks)
        try:
            dp = (getattr(self.cfg.solidworks, 'description_prop', 'DESCRIZIONE') or 'DESCRIZIONE').strip()
            if dp and doc.description:
                props[str(dp)] = str(doc.description)
        except Exception:
            pass

        created_model = False
        created_drw = False

        if need_model:
            r1 = create_model_file(sw, tpl_model, str(out_model), props=props)
            if not r1.ok:
                warn(r1.message + ("\n\n" + r1.details if r1.details else ""))
                return
            created_model = True

        if need_drw:
            r2 = create_drawing_file(sw, self.cfg.solidworks.template_drawing, str(out_drw), props=props)
            if r2.ok:
                created_drw = True
            else:
                warn(r2.message + ("\n\n" + r2.details if r2.details else ""))

        up: dict[str, str] = {}
        if out_model.is_file() or created_model:
            up["file_wip_path"] = str(out_model)
        if out_drw.is_file() or created_drw:
            up["file_wip_drw_path"] = str(out_drw)
        if up:
            self.store.update_document(doc.code, **up)

        try:
            if out_model.is_file() or created_model:
                doc_sync = self.store.get_document(doc.code) or doc
                if bool(getattr(doc_sync, "checked_out", False)) and self._is_doc_checked_out_by_me(doc_sync):
                    self._sync_sw_to_pdm(doc.code)
        except Exception:
            pass

        if only_missing:
            made = []
            if created_model:
                made.append("MODEL")
            if created_drw:
                made.append("DRW")
            if made:
                info(f"Creati file mancanti: {', '.join(made)}.")
            else:
                info("Nessun nuovo file creato.")
        else:
            info("Creazione file completata.")
        self._refresh_all()

    # ---------------- Tab: Gerarchia
    # ESTRATTO IN: pdm_sw/gui/tab_gerarchia.py -> TabGerarchia
    # def _ui_gerarchia(self):
    #     tab = self.tab_gerarchia
    #
    #     actions = ctk.CTkFrame(tab, fg_color="#F2D65C")
    #     actions.pack(fill="x", padx=10, pady=(10, 6))
    #
    #     ctk.CTkButton(actions, text="AGGIORNA", width=120, command=self._refresh_hierarchy_tree).pack(side="left", padx=6, pady=6)
    #     ctk.CTkButton(actions, text="REPORT GENERALE", width=160, command=self._generate_hierarchy_report).pack(side="left", padx=6, pady=6)
    #     self.hierarchy_include_obs_var = ctk.BooleanVar(value=False)
    #     ctk.CTkCheckBox(
    #         actions,
    #         text="Mostra OBS",
    #         variable=self.hierarchy_include_obs_var,
    #         command=self._refresh_hierarchy_tree,
    #     ).pack(side="left", padx=(12, 6), pady=6)
    #
    #     ctk.CTkLabel(
    #         actions,
    #         text="Struttura: MMM -> GGGG -> codici PART/ASSY",
    #     ).pack(side="right", padx=10, pady=6)
    #
    #     tree_frame = ctk.CTkFrame(tab)
    #     tree_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    #
    #     style = ttk.Style()
    #     try:
    #         base = tkfont.nametofont("TkDefaultFont")
    #         base_size = int(base.cget("size"))
    #         tree_size = max(10, int(round(base_size * 1.45)))
    #         tree_family = base.cget("family")
    #     except Exception:
    #         tree_size = 15
    #         tree_family = "Segoe UI"
    #     style.configure(
    #         "PDM.Hierarchy.Treeview",
    #         font=(tree_family, tree_size),
    #         rowheight=int(tree_size * 2),
    #     )
    #
    #     self.hierarchy_tree = ttk.Treeview(tree_frame, show="tree", style="PDM.Hierarchy.Treeview")
    #     self.hierarchy_tree.tag_configure("part_node", foreground="#0B5ED7")
    #     self.hierarchy_tree.tag_configure("assy_node", foreground="#2E7D32")
    #     self.hierarchy_tree.tag_configure("machine_node", foreground="#E67E22", font=("Segoe UI", tree_size, "bold"))
    #     self.hierarchy_tree.tag_configure("group_node", foreground="#3498DB", font=("Segoe UI", tree_size, "bold"))
    #     yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.hierarchy_tree.yview)
    #     xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.hierarchy_tree.xview)
    #     self.hierarchy_tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
    #
    #     self.hierarchy_tree.grid(row=0, column=0, sticky="nsew")
    #     yscroll.grid(row=0, column=1, sticky="ns")
    #     xscroll.grid(row=1, column=0, sticky="ew")
    #     tree_frame.grid_rowconfigure(0, weight=1)
    #     tree_frame.grid_columnconfigure(0, weight=1)
    #
    #     self.hierarchy_tree.bind("<Double-1>", self._on_hierarchy_double_click)
    #
    # def _assoc_path_status(self, path: str) -> str:
    #     p = (path or "").strip()
    #     if not p:
    #         return "NON_ASSOCIATO"
    #     return "OK" if Path(p).exists() else "MANCANTE"
    #
    # def _refresh_hierarchy_tree(self):
    #     if not hasattr(self, "hierarchy_tree"):
    #         return
    #
    #     tree = self.hierarchy_tree
    #     current_nodes = tree.get_children()
    #     if current_nodes:
    #         tree.delete(*current_nodes)
    #
    #     include_obs = bool(self.hierarchy_include_obs_var.get()) if hasattr(self, "hierarchy_include_obs_var") else False
    #
    #     machines_raw = self.store.list_machines()
    #     machine_names: dict[str, str] = {mmm: name for mmm, name in machines_raw}
    #
    #     groups_by_machine: dict[str, dict[str, str]] = {}
    #     for mmm, _ in machines_raw:
    #         groups_by_machine[mmm] = {gggg: g_name for gggg, g_name in self.store.list_groups(mmm)}
    #
    #     docs = self.store.list_documents(include_obs=include_obs)
    #     docs_by_pair: dict[tuple[str, str], list[Document]] = defaultdict(list)
    #     docs_machine: dict[str, list[Document]] = defaultdict(list)  # MACHINE per MMM
    #
    #     for d in docs:
    #         if d.doc_type == "MACHINE":
    #             docs_machine[d.mmm].append(d)
    #         else:
    #             docs_by_pair[(d.mmm, d.gggg)].append(d)
    #         machine_names.setdefault(d.mmm, "")
    #         groups_by_machine.setdefault(d.mmm, {})
    #         if d.gggg:
    #             groups_by_machine[d.mmm].setdefault(d.gggg, "")
    #
    #     if not machine_names:
    #         tree.insert("", "end", text="(nessun MMM disponibile)")
    #         return
    #
    #     for mmm in sorted(machine_names.keys()):
    #         m_name = (machine_names.get(mmm) or "").strip()
    #         m_text = f"{mmm} - {m_name}" if m_name else mmm
    #         m_node = tree.insert("", "end", text=m_text, open=True)
    #
    #         # Mostra MACHINE versions sotto MMM
    #         machine_docs = docs_machine.get(mmm, [])
    #         if machine_docs:
    #             machine_docs.sort(key=lambda d: d.code)
    #             for d in machine_docs:
    #                 model_path = d.best_model_path_for_state()
    #                 drw_path = d.best_drw_path_for_state()
    #                 d_node = tree.insert(m_node, "end", text=f"📦 {d.code}", values=(d.code,), tags=("machine_node",))
    #                 tree.insert(
    #                     d_node,
    #                     "end",
    #                     text=f"DESC: {d.description}",
    #                 )
    #                 tree.insert(
    #                     d_node,
    #                     "end",
    #                     text=f"MODEL ({d.state}): {model_path if model_path else 'NON ASSOCIATO'}",
    #                 )
    #                 tree.insert(
    #                     d_node,
    #                     "end",
    #                     text=f"DRW ({d.state}): {drw_path if drw_path else 'NON ASSOCIATO'}",
    #                 )
    #
    #         group_map = groups_by_machine.get(mmm, {})
    #         gggg_keys = sorted(group_map.keys())
    #         if not gggg_keys:
    #             if not machine_docs:
    #                 tree.insert(m_node, "end", text="(nessun GGGG o versione macchina)")
    #             continue
    #
    #         for gggg in gggg_keys:
    #             g_name = (group_map.get(gggg) or "").strip()
    #             dlist = docs_by_pair.get((mmm, gggg), [])
    #             dlist.sort(key=lambda d: (0 if d.doc_type == "PART" else (1 if d.doc_type == "ASSY" else 2), d.code))
    #
    #             part_count = sum(1 for d in dlist if d.doc_type == "PART")
    #             assy_count = sum(1 for d in dlist if d.doc_type == "ASSY")
    #             group_count = sum(1 for d in dlist if d.doc_type == "GROUP")
    #
    #             g_base = f"{gggg} - {g_name}" if g_name else gggg
    #             g_text = f"{g_base} (GROUP:{group_count} PART:{part_count} ASSY:{assy_count})"
    #             g_node = tree.insert(m_node, "end", text=g_text, open=True)
    #
    #             if not dlist:
    #                 tree.insert(g_node, "end", text="(nessun codice)")
    #                 continue
    #
    #             for d in dlist:
    #                 model_path = d.best_model_path_for_state()
    #                 drw_path = d.best_drw_path_for_state()
    #                 if d.doc_type == "PART":
    #                     tags = ("part_node",)
    #                     icon = ""
    #                 elif d.doc_type == "ASSY":
    #                     tags = ("assy_node",)
    #                     icon = ""
    #                 elif d.doc_type == "GROUP":
    #                     tags = ("group_node",)
    #                     icon = "📁 "
    #                 else:
    #                     tags = ()
    #                     icon = ""
    #                 d_node = tree.insert(g_node, "end", text=f"{icon}{d.code}", values=(d.code,), tags=tags)
    #                 if d.doc_type == "GROUP":
    #                     tree.insert(
    #                         d_node,
    #                         "end",
    #                         text=f"DESC: {d.description}",
    #                     )
    #                 tree.insert(
    #                     d_node,
    #                     "end",
    #                     text=f"MODEL ({d.state}): {model_path if model_path else 'NON ASSOCIATO'}",
    #                 )
    #                 tree.insert(
    #                     d_node,
    #                     "end",
    #                     text=f"DRW ({d.state}): {drw_path if drw_path else 'NON ASSOCIATO'}",
    #                 )
    #
    # def _on_hierarchy_double_click(self, _evt=None):
    #     if not hasattr(self, "hierarchy_tree"):
    #         return
    #     try:
    #         sel = self.hierarchy_tree.selection()
    #         if not sel:
    #             return
    #         values = self.hierarchy_tree.item(sel[0]).get("values", [])
    #         code = str(values[0]).strip() if values else ""
    #         if not code:
    #             return
    #         self.wf_code_var.set(code)
    #         try:
    #             self.tabs.set("Operativo")
    #         except Exception:
    #             pass
    #         self._refresh_workflow_panel()
    #     except Exception:
    #         return

    # ---------------- Tab: Consultazione
    def _select_doc_by_code(self, code: str):
        self.selected_code = code
        self._refresh_workflow_panel()

    def _get_table_selected_code(self) -> str:
        """Ritorna il codice selezionato dalla tabella attiva (preferisce Ricerca&Consultazione)."""
        # Tab Operativo (fonte primaria): usa helper del tab se disponibile
        try:
            t = getattr(self, "tab_operativo_obj", None)
            if t is not None and hasattr(t, "_get_selected_rc_code"):
                code = str(t._get_selected_rc_code() or "").strip()
                if code:
                    return code
        except Exception:
            pass

        # Nuova tab unificata
        if hasattr(self, "rc_table") and getattr(self, "rc_table", None) is not None:
            try:
                sel = self.rc_table.tree.selection()
                if not sel:
                    fid = self.rc_table.tree.focus()
                    if fid:
                        sel = (fid,)
                if not sel:
                    return ""
                vals = self.rc_table.tree.item(sel[0]).get("values", [])
                idx = getattr(self.rc_table, "key_index", 2)
                return str(vals[idx]) if vals and len(vals) > idx else (str(vals[0]) if vals else "")
            except Exception:
                pass

        # Vecchia consultazione (fallback)
        if hasattr(self, "docs_table") and getattr(self, "docs_table", None) is not None:
            try:
                sel = self.docs_table.tree.selection()
                if not sel:
                    return ""
                vals = self.docs_table.tree.item(sel[0]).get("values", [])
                idx = getattr(self.docs_table, "key_index", 0)
                return str(vals[idx]) if vals and len(vals) > idx else (str(vals[0]) if vals else "")
            except Exception:
                pass

        # Vecchia ricerca (fallback)
        if hasattr(self, "search_table") and getattr(self, "search_table", None) is not None:
            try:
                sel = self.search_table.tree.selection()
                if not sel:
                    return ""
                vals = self.search_table.tree.item(sel[0]).get("values", [])
                idx = getattr(self.search_table, "key_index", 0)
                return str(vals[idx]) if vals and len(vals) > idx else (str(vals[0]) if vals else "")
            except Exception:
                pass

        return ""



    # --- Search -> Send to Workflow helpers ---



    def _build_table_schema_with_sw_props(self):
        """Costruisce schema tabella (Consultazione/Ricerca) includendo colonne dinamiche per proprieta SW lette.

        Ritorna: (columns, headings, props, key_index)
        - columns: nomi colonne Treeview
        - headings: testi intestazioni
        - props: lista nomi proprieta SW (uppercase) da leggere dal DB (doc_custom_values)
        - key_index: indice della colonna 'code' (chiave univoca)
        """
        base_cols = ["m_ok", "d_ok", "code", "doc_type", "revision", "state", "check_state", "description"]
        base_heads = ["M", "D", "CODICE", "TIPO", "REV", "STATO", "CHECK", "DESCRIZIONE"]

        # proprieta SW da leggere configurate in Tab SolidWorks
        props = []
        try:
            props = list(getattr(self.cfg.solidworks, "read_properties", []) or [])
        except Exception:
            props = []
        # normalizza: uppercase + unici preservando ordine
        seen = set()
        props_u = []
        for p in props:
            pn = str(p).strip().upper()
            if not pn:
                continue
            if pn in seen:
                continue
            seen.add(pn)
            props_u.append(pn)

        columns = list(base_cols) + [f"sw_{p}" for p in props_u]
        headings = list(base_heads) + props_u
        key_index = 2  # code e terza colonna (dopo M,D)
        return columns, headings, props_u, key_index

    def _flag_file_exists(self, path_s: str) -> str:
        return "OK" if (path_s and Path(path_s).is_file()) else ""

    def _state_row_tag(self, state: str) -> str:
        s = (state or "").strip().upper()
        return {
            "WIP": "state_wip",
            "IN_REV": "state_in_rev",
            "REL": "state_rel",
            "OBS": "state_obs",
        }.get(s, "")

    def _best_model_and_drw_paths(self, doc: Document) -> tuple[str, str]:
        model_path_s = ""
        drw_path_s = ""
        try:
            model_path_s = doc.best_model_path_for_state() or ""
        except Exception:
            try:
                model_path_s = doc.best_path_for_state() or ""
            except Exception:
                model_path_s = ""
        try:
            drw_path_s = doc.best_drw_path_for_state() or ""
        except Exception:
            try:
                drw_path_s = doc.best_drawing_path_for_state() or ""
            except Exception:
                drw_path_s = ""
        return model_path_s, drw_path_s

    def _model_and_drawing_flags(self, doc: Document) -> tuple[str, str]:
        model_path_s, drw_path_s = self._best_model_and_drw_paths(doc)
        return self._flag_file_exists(model_path_s), self._flag_file_exists(drw_path_s)

    def _list_rev_files_for_doc(self, doc: Document) -> tuple[list[str], list[str]]:
        model_rev_files: list[str] = []
        drw_rev_files: list[str] = []
        root = str(getattr(self.cfg.solidworks, "archive_root", "") or "").strip()
        if not root:
            return model_rev_files, drw_rev_files
        try:
            _wip, _rel, _inrev, rev = archive_dirs(root, doc.mmm, doc.gggg)
        except Exception:
            return model_rev_files, drw_rev_files

        model_ext = ".sldprt" if str(doc.doc_type).upper() == "PART" else ".sldasm"
        model_pat = f"{doc.code}_R*{model_ext}"
        drw_pat = f"{doc.code}_R*.slddrw"

        try:
            model_rev_files = [str(p) for p in sorted(rev.glob(model_pat), key=lambda x: x.name, reverse=True) if p.is_file()]
        except Exception:
            model_rev_files = []
        try:
            drw_rev_files = [str(p) for p in sorted(rev.glob(drw_pat), key=lambda x: x.name, reverse=True) if p.is_file()]
        except Exception:
            drw_rev_files = []

        return model_rev_files, drw_rev_files

    def _refresh_docs_table_DEPRECATED(self):
        """Aggiorna la tabella Consultazione (documenti) includendo colonne SW (custom lette)."""
        try:
            include_obs = bool(self.include_obs_var.get()) if hasattr(self, "include_obs_var") else False
        except Exception:
            include_obs = False

        docs = self.store.list_documents(include_obs=include_obs)

        columns, headings, props, key_index = self._build_table_schema_with_sw_props()
        # aggiorna schema tabella se necessario
        if hasattr(self, "docs_table") and self.docs_table is not None:
            try:
                self.docs_table.set_schema(columns, headings, key_index=key_index)
            except Exception:
                pass

        codes = [d.code for d in docs]
        custom_bulk = self.store.get_custom_values_bulk(codes, props) if props else {}

        rows = []
        for d in docs:
            try:
                rev = f"{int(d.revision):02d}"
            except Exception:
                rev = str(d.revision)
            m_ok, d_ok = self._model_and_drawing_flags(d)
            vals = []
            if props:
                cdict = custom_bulk.get(d.code, {})
                for pn in props:
                    vals.append(cdict.get(pn, ""))
            row_values = [m_ok, d_ok, d.code, d.doc_type, rev, d.state, self._checkout_table_value(d), d.description, *vals]
            row_tag = self._state_row_tag(d.state)
            rows.append({"values": row_values, "tags": (row_tag,) if row_tag else ()})

        if hasattr(self, "docs_table") and self.docs_table is not None:
            self.docs_table.set_rows(rows)

    def _is_solidworks_process_running(self) -> bool:
        """Verifica leggera processo SolidWorks senza aprire istanze COM."""
        try:
            cp = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq SLDWORKS.exe"],
                capture_output=True,
                text=True,
                check=False,
            )
            out = f"{cp.stdout}\n{cp.stderr}".casefold()
            return "sldworks.exe" in out
        except Exception:
            return False

    def _apply_sldreg_before_sw_launch(self, *, code: str, open_source: str, kind: str) -> None:
        """Applica import .sldreg filtrato solo prima dell'avvio di SolidWorks."""
        sw_cfg = getattr(self.cfg, "solidworks", None)
        if sw_cfg is None:
            return
        if not bool(getattr(sw_cfg, "sldreg_enabled", False)):
            return

        sldreg_file = str(getattr(sw_cfg, "sldreg_file", "") or "").strip()
        cleanup = bool(getattr(sw_cfg, "sldreg_cleanup_before_import", True))
        restore_options = SldregRestoreOptions(
            system_options=bool(getattr(sw_cfg, "sldreg_restore_system_options", True)),
            toolbar_layout=bool(getattr(sw_cfg, "sldreg_restore_toolbar_layout", True)),
            toolbar_mode=str(getattr(sw_cfg, "sldreg_restore_toolbar_mode", "all")),
            keyboard_shortcuts=bool(getattr(sw_cfg, "sldreg_restore_keyboard_shortcuts", True)),
            mouse_gestures=bool(getattr(sw_cfg, "sldreg_restore_mouse_gestures", True)),
            menu_customizations=bool(getattr(sw_cfg, "sldreg_restore_menu_customizations", True)),
            saved_views=bool(getattr(sw_cfg, "sldreg_restore_saved_views", True)),
        )

        if not sldreg_file:
            self._log_activity(
                action="SW_SLDREG",
                code=code,
                status="WARN",
                message="Gestione .sldreg attiva ma file non configurato. Avvio pulito.",
                details={"open_source": open_source, "kind": kind},
            )
            return

        source = Path(sldreg_file)
        if not source.exists():
            self._log_activity(
                action="SW_SLDREG",
                code=code,
                status="WARN",
                message=f"File .sldreg non trovato: {source}",
                details={"open_source": open_source, "kind": kind, "path": str(source)},
            )
            return

        if self._is_solidworks_process_running():
            self._log_activity(
                action="SW_SLDREG",
                code=code,
                status="SKIP",
                message="Import .sldreg filtrato saltato: SolidWorks gia avviato.",
                details={"open_source": open_source, "kind": kind, "path": str(source)},
            )
            return

        ok, msg = import_sldreg_filtered(
            source,
            cleanup_before_import=cleanup,
            restore_options=restore_options,
        )
        self._log_activity(
            action="SW_SLDREG",
            code=code,
            status="OK" if ok else "WARN",
            message=msg if ok else f"Import .sldreg non applicato: {msg}. Avvio pulito.",
            details={
                "open_source": open_source,
                "kind": kind,
                "path": str(source),
                "cleanup_before_import": cleanup,
                "restore_options": {
                    "system_options": bool(restore_options.system_options),
                    "toolbar_layout": bool(restore_options.toolbar_layout),
                    "toolbar_mode": str(restore_options.toolbar_mode),
                    "keyboard_shortcuts": bool(restore_options.keyboard_shortcuts),
                    "mouse_gestures": bool(restore_options.mouse_gestures),
                    "menu_customizations": bool(restore_options.menu_customizations),
                    "saved_views": bool(restore_options.saved_views),
                },
            },
        )

    def _open_path_in_solidworks(self, *, code: str, file_path: Path, open_source: str, kind: str) -> bool:
        """Apre un file CAD in SolidWorks.

        Tenta prima apertura COM (avvia SW se chiuso), poi fallback shell.
        """
        p = Path(file_path)
        if not p.exists():
            self._log_activity(
                action="OPEN_FILE",
                code=code,
                status="WARN",
                message=f"{kind} mancante: {p.name}",
                details={"path": str(p), "open_source": open_source, "kind": kind},
            )
            return False

        self._apply_sldreg_before_sw_launch(code=code, open_source=open_source, kind=kind)

        com_err = ""
        try:
            sw, res = get_solidworks_app(visible=True, timeout_s=25.0, allow_launch=True)
            if sw is not None and res.ok:
                opened = open_doc(sw, str(p), silent=False)
                if opened is not None:
                    self._log_activity(
                        action="OPEN_FILE",
                        code=code,
                        status="OK",
                        message=f"Aperto {kind} {p.name}",
                        details={"path": str(p), "open_source": open_source, "kind": kind, "open_mode": "COM"},
                    )
                    return True
                com_err = "OpenDoc6 ha restituito None."
            else:
                com_err = str(getattr(res, "message", "") or "SolidWorks non disponibile via COM.")
        except Exception as e:
            com_err = str(e)

        try:
            os.startfile(str(p))  # type: ignore[attr-defined]
            self._log_activity(
                action="OPEN_FILE",
                code=code,
                status="OK",
                message=f"Aperto {kind} {p.name}",
                details={
                    "path": str(p),
                    "open_source": open_source,
                    "kind": kind,
                    "open_mode": "SHELL",
                    "com_error": com_err,
                },
            )
            return True
        except Exception as e:
            self._log_activity(
                action="OPEN_FILE",
                code=code,
                status="ERROR",
                message=f"Errore apertura {kind} {p.name}",
                details={
                    "path": str(p),
                    "open_source": open_source,
                    "kind": kind,
                    "open_mode": "FAILED",
                    "com_error": com_err,
                    "shell_error": str(e),
                },
            )
            return False


    def _doc_dbl(self, code: str):
        """Doppio click su riga elenco: apre prima MODELLO, poi fallback DRW."""
        try:
            code = str(code).strip()
        except Exception:
            code = ""
        if not code:
            return
        doc = self.store.get_document(code)
        if not doc:
            return

        model_path = ""
        try:
            model_path = doc.best_model_path_for_state()
        except Exception:
            try:
                model_path = doc.best_path_for_state()
            except Exception:
                model_path = ""
        if model_path:
            p = Path(model_path)
            if p.exists():
                try:
                    if self._open_path_in_solidworks(
                        code=doc.code,
                        file_path=p,
                        open_source="RC_DOUBLE_CLICK",
                        kind="MODEL",
                    ):
                        return
                except Exception:
                    pass

        drw_path = ""
        try:
            drw_path = doc.best_drw_path_for_state()
        except Exception:
            try:
                drw_path = doc.best_drawing_path_for_state()
            except Exception:
                drw_path = ""
        if drw_path:
            p = Path(drw_path)
            if p.exists():
                try:
                    if self._open_path_in_solidworks(
                        code=doc.code,
                        file_path=p,
                        open_source="RC_DOUBLE_CLICK",
                        kind="DRW",
                    ):
                        return
                except Exception as e:
                    self._log_activity(
                        action="OPEN_FILE",
                        code=doc.code,
                        status="ERROR",
                        message=f"Doppio click: errore apertura DRW {p.name}",
                        details={"path": str(p), "open_source": "RC_DOUBLE_CLICK", "kind": "DRW"},
                    )
                    warn(f"Impossibile aprire il file:\n{p}\n\n{e}")
                    return

        self._log_activity(
            action="OPEN_FILE",
            code=doc.code,
            status="WARN",
            message="Doppio click: nessun file disponibile",
            details={"open_source": "RC_DOUBLE_CLICK"},
        )
        warn("Nessun file disponibile (MODEL/DRW) per questo codice/stato.")


    def _open_selected_model(self):
        code = self._get_table_selected_code()
        if not code:
            warn("Seleziona un codice nell'elenco Operativo.")
            return
        doc = self.store.get_document(code)
        if not doc:
            warn("Documento non trovato.")
            return
        candidates = []
        try:
            candidates.append(str(doc.best_model_path_for_state() or ""))
        except Exception:
            pass
        candidates.extend([
            str(getattr(doc, "file_wip_path", "") or ""),
            str(getattr(doc, "file_inrev_path", "") or ""),
            str(getattr(doc, "file_rel_path", "") or ""),
        ])
        seen = set()
        ordered = []
        for c in candidates:
            p = str(c or "").strip()
            if not p:
                continue
            k = p.lower()
            if k in seen:
                continue
            seen.add(k)
            ordered.append(p)
        path = ""
        for cand in ordered:
            if Path(cand).exists():
                path = cand
                break
        if not path and ordered:
            path = ordered[0]
        if not path:
            warn("Percorso file non disponibile per questo stato.")
            return
        p = Path(path)
        if not p.exists():
            self._log_activity(
                action="OPEN_FILE",
                code=doc.code,
                status="WARN",
                message=f"MODEL mancante: {p.name}",
                details={"path": str(p), "open_source": "RC_BUTTON_MODEL", "kind": "MODEL"},
            )
            warn(f"File non trovato:\n{p}")
            return
        if not self._open_path_in_solidworks(
            code=doc.code,
            file_path=p,
            open_source="RC_BUTTON_MODEL",
            kind="MODEL",
        ):
            warn(f"Impossibile aprire il file:\n{p}\n\nVerifica installazione/avvio SolidWorks.")

    def _open_selected_drw(self):
        code = self._get_table_selected_code()
        if not code:
            warn("Seleziona un codice nell'elenco Operativo.")
            return
        doc = self.store.get_document(code)
        if not doc:
            warn("Documento non trovato.")
            return
        candidates = []
        try:
            candidates.append(str(doc.best_drw_path_for_state() or ""))
        except Exception:
            pass
        candidates.extend([
            str(getattr(doc, "file_wip_drw_path", "") or ""),
            str(getattr(doc, "file_inrev_drw_path", "") or ""),
            str(getattr(doc, "file_rel_drw_path", "") or ""),
        ])
        seen = set()
        ordered = []
        for c in candidates:
            p = str(c or "").strip()
            if not p:
                continue
            k = p.lower()
            if k in seen:
                continue
            seen.add(k)
            ordered.append(p)
        path = ""
        for cand in ordered:
            if Path(cand).exists():
                path = cand
                break
        if not path and ordered:
            path = ordered[0]
        if not path:
            self._log_activity(
                action="OPEN_FILE",
                code=doc.code,
                status="WARN",
                message="DRW non disponibile",
                details={"open_source": "RC_BUTTON_DRW", "kind": "DRW"},
            )
            warn("Disegno non disponibile per questo codice/stato.")
            return
        p = Path(path)
        if not p.exists():
            self._log_activity(
                action="OPEN_FILE",
                code=doc.code,
                status="WARN",
                message=f"DRW mancante: {p.name}",
                details={"path": str(p), "open_source": "RC_BUTTON_DRW", "kind": "DRW"},
            )
            warn(f"Disegno non trovato:\n{p}")
            return
        if not self._open_path_in_solidworks(
            code=doc.code,
            file_path=p,
            open_source="RC_BUTTON_DRW",
            kind="DRW",
        ):
            warn(f"Impossibile aprire il disegno:\n{p}\n\nVerifica installazione/avvio SolidWorks.")

    # ---------------- Sync proprieta (PDM -> SolidWorks)
    def _sync_pdm_to_sw(self, code: str):
        """Scrive le proprieta CORE (mappatura PDM->SW) nel file SolidWorks esistente."""
        doc = self.store.get_document(code)
        if not doc:
            return False
        if not self._require_checkout_for_edit(doc, "Sync PDM->SW"):
            return False
        model_path_s = ""
        try:
            model_path_s = doc.best_model_path_for_state()
        except Exception:
            model_path_s = doc.file_rel_path or doc.file_inrev_path or doc.file_wip_path
        if not model_path_s:
            warn("Nessun file modello associato al codice.")
            return False

        from pathlib import Path
        from pdm_sw.archive import set_readonly
        from pdm_sw.sw_api import open_doc, set_custom_properties, save_existing_doc, close_doc, get_solidworks_app

        self._apply_sldreg_before_sw_launch(code=doc.code, open_source="SYNC_PDM_TO_SW", kind="MODEL")
        sw, res = get_solidworks_app(visible=False, timeout_s=30.0)
        if not res.ok or sw is None:
            warn(res.message + ("\n\n" + res.details if res.details else ""))
            return False

        final_ro = (doc.state not in ("WIP", "IN_REV"))
        try:
            # rendi scrivibile per salvare le proprieta
            set_readonly(Path(model_path_s), False)
        except Exception:
            pass

        try:
            mdl = open_doc(sw, model_path_s, silent=True)
            if mdl is None:
                warn("Impossibile aprire il file in SolidWorks.")
                return False
            props = self._build_sw_props_for_doc(doc)  # CORE only
            set_custom_properties(mdl, props)
            save_existing_doc(mdl)
            close_doc(sw, mdl)
            return True
        except Exception as e:
            warn(f"Sync PDM->SW fallita: {e}")
            return False
        finally:
            try:
                set_readonly(Path(model_path_s), final_ro)
            except Exception:
                pass

    def _sync_sw_to_pdm(self, code: str):
        """Legge descrizione + proprieta custom (configurate) da SolidWorks e aggiorna PDM."""
        doc = self.store.get_document(code)
        if not doc:
            return False
        if not self._require_checkout_for_edit(doc, "Sync SW->PDM"):
            return False
        model_path_s = ""
        try:
            model_path_s = doc.best_model_path_for_state()
        except Exception:
            model_path_s = doc.file_rel_path or doc.file_inrev_path or doc.file_wip_path
        if not model_path_s:
            warn("Nessun file modello associato al codice.")
            return False

        from pdm_sw.sw_api import open_doc, get_custom_properties, close_doc, get_solidworks_app

        self._apply_sldreg_before_sw_launch(code=doc.code, open_source="SYNC_SW_TO_PDM", kind="MODEL")
        sw, res = get_solidworks_app(visible=False, timeout_s=30.0)
        if not res.ok or sw is None:
            warn(res.message + ("\n\n" + res.details if res.details else ""))
            return False

        try:
            mdl = open_doc(sw, model_path_s, silent=True)
            if mdl is None:
                warn("Impossibile aprire il file in SolidWorks.")
                return False
            props = get_custom_properties(mdl) or {}
            up = {str(k).strip().upper(): str(v) for k, v in props.items()}

            desc_prop = (getattr(self.cfg.solidworks, "description_prop", "DESCRIZIONE") or "DESCRIZIONE").strip().upper()
            sw_desc = up.get(desc_prop, "").strip()
            if sw_desc:
                self.store.update_document(code, description=sw_desc)

            try:
                rp = list(getattr(self.cfg.solidworks, "read_properties", []) or [])
            except Exception:
                rp = []
            for p in rp:
                pn = str(p).strip().upper()
                if not pn or pn == desc_prop:
                    continue
                val = up.get(pn, "")
                self.store.set_custom_value(code, pn, val)

            close_doc(sw, mdl)
            return True
        except Exception as e:
            warn(f"Sync SW->PDM fallita: {e}")
            return False

    def _force_pdm_to_sw_selected(self):
        code = self._get_workflow_or_selected_code()
        if not code:
            warn("Seleziona un codice in Consultazione o nel pannello Workflow.")
            return
        if self._sync_pdm_to_sw(code):
            info("Sync PDM->SolidWorks completata (best-effort).")

    def _force_sw_to_pdm_selected(self):
        code = self._get_workflow_or_selected_code()
        if not code:
            warn("Seleziona un codice in Consultazione o nel pannello Workflow.")
            return
        if self._sync_sw_to_pdm(code):
            info("Lettura SolidWorks->PDM completata (best-effort).")

    def _create_files_for_selected(self):
        code = self._get_table_selected_code()
        if not code:
            warn("Seleziona un codice in Consultazione.")
            return
        self._create_files_for_code(code)

    def _create_missing_files_for_selected(self):
        code = self._get_table_selected_code()
        if not code:
            warn("Seleziona un codice in Consultazione.")
            return
        self._create_files_for_code(code, create_drw=True, only_missing=True)

    def _get_search_selected_code(self) -> str:
        if not hasattr(self, "search_table"):
            return ""
        sel = self.search_table.tree.selection()
        if not sel:
            return ""
        vals = self.search_table.tree.item(sel[0]).get("values", [])
        try:
            idx = getattr(self.search_table, 'key_index', 0)
        except Exception:
            idx = 0
        return str(vals[idx]) if vals and len(vals) > idx else (str(vals[0]) if vals else "")

    def _on_search_selection(self, _evt=None):
        code = self._get_search_selected_code()
        if hasattr(self, "btn_send_to_wf"):
            self.btn_send_to_wf.configure(state="normal" if code else "disabled")

    def _send_search_to_workflow(self):
        code = self._get_search_selected_code()
        if not code:
            return
        # set workflow selected code and navigate to workflow tab
        if hasattr(self, "wf_code_var"):
            self.wf_code_var.set(code)
        try:
            self.tabs.set("Operativo")
        except Exception:
            pass
        try:
            self._refresh_workflow_panel()
        except Exception:
            pass


    def _run_search(self):
        """(DEPRECATO) Vecchia tab Ricerca rimossa. Usa tab Operativo."""
        return

    # ===============================================================================
    # TAB OPERATIVO - ESTRATTO IN: pdm_sw/gui/tab_operativo.py -> TabOperativo
    # Include: paned window, ricerca/consultazione, workflow panel
    # Metodi estratti: _ui_operativo(), _apply_operativo_split(), 
    # _on_operativo_sash_move(), _ui_ricerca_consultazione(), _on_rc_select(),
    # _get_selected_rc_code(), _send_rc_to_workflow(), _reset_rc_filters(),
    # _search_rc(), _refresh_rc_table(), _ui_workflow(), _load_selected_doc(),
    # _update_workflow_buttons(), _refresh_workflow_panel()
    # ===============================================================================

    def _wf_backup_event(self, reason: str):
        if not self.cfg.backup.enabled:
            return
        res = self.backup.backup_now(reason, force=True)
        if not res.ok:
            warn(res.message)

    def _save_workflow_doc(self, doc: Document) -> None:
        self.store.update_document(
            doc.code,
            state=doc.state,
            revision=doc.revision,
            obs_prev_state=getattr(doc, "obs_prev_state", "") or "",
            file_wip_path=doc.file_wip_path,
            file_rel_path=doc.file_rel_path,
            file_inrev_path=doc.file_inrev_path,
            file_wip_drw_path=doc.file_wip_drw_path,
            file_rel_drw_path=doc.file_rel_drw_path,
            file_inrev_drw_path=doc.file_inrev_drw_path,
        )

    def _apply_operativo_split(self) -> None:
        paned = getattr(self, "_operativo_paned", None)
        if paned is None:
            return
        ratio = self._clamp_workflow_width_ratio(self.workflow_width_ratio)
        self.workflow_width_ratio = ratio
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
        paned = getattr(self, "_operativo_paned", None)
        if paned is None:
            return
        try:
            total = int(paned.winfo_width() or 0)
            if total <= 20:
                return
            sash_x = int(paned.sashpos(0))
            right_ratio = 1.0 - (float(sash_x) / float(total))
            self.workflow_width_ratio = self._clamp_workflow_width_ratio(right_ratio)
            if save:
                self._save_local_settings()
        except Exception:
            pass

    def _ui_ricerca_consultazione(self):
        """Tab unica: Ricerca&Consultazione.
        Barra superiore: azioni (gialla).
        Barra inferiore: filtri + pulsanti ricerca.
        """
        tab = self.tab_rc

        # --- barra azioni (gialla) ---
        actions = ctk.CTkFrame(tab, fg_color="#F2D65C")  # giallo tenue
        actions.pack(fill="x", padx=10, pady=(10, 6))

        # pulsanti (ex Consultazione)
        ctk.CTkButton(actions, text="AGGIORNA", width=120, command=self._refresh_rc_table).pack(side="left", padx=6, pady=6)
        ctk.CTkButton(actions, text="APRI MODELLO", width=140, command=self._open_selected_model).pack(side="left", padx=6, pady=6)
        ctk.CTkButton(actions, text="APRI DRW", width=120, command=self._open_selected_drw).pack(side="left", padx=6, pady=6)
        ctk.CTkButton(actions, text="CREA FILE MANCANTI", width=170, command=self._create_missing_files_for_selected).pack(side="left", padx=6, pady=6)
        ctk.CTkButton(actions, text="COPIA CODICE", width=130, command=self._copy_selected_code_to_new_wip).pack(side="left", padx=6, pady=6)
        ctk.CTkButton(actions, text="FORZA PDM->SW", width=150, command=self._force_pdm_to_sw_selected).pack(side="left", padx=6, pady=6)
        ctk.CTkButton(actions, text="FORZA SW->PDM", width=150, command=self._force_sw_to_pdm_selected).pack(side="left", padx=6, pady=6)

        # --- barra filtri ---
        filters = ctk.CTkFrame(tab)
        filters.pack(fill="x", padx=10, pady=(0, 8))

        # testo libero
        ctk.CTkLabel(filters, text="Testo:").pack(side="left", padx=(6, 2), pady=6)
        self.search_text_var = ctk.StringVar(value="")
        ctk.CTkEntry(filters, textvariable=self.search_text_var, width=220).pack(side="left", padx=4, pady=6)

        # Stato
        ctk.CTkLabel(filters, text="Stato:").pack(side="left", padx=(10, 2), pady=6)
        self.search_state_var = ctk.StringVar(value="")
        ctk.CTkComboBox(filters, variable=self.search_state_var, values=["", "WIP", "REL", "IN_REV", "OBS"], width=110).pack(side="left", padx=4, pady=6)

        # Tipo
        ctk.CTkLabel(filters, text="Tipo:").pack(side="left", padx=(10, 2), pady=6)
        self.search_type_var = ctk.StringVar(value="")
        ctk.CTkComboBox(filters, variable=self.search_type_var, values=["", "PART", "ASSY"], width=110).pack(side="left", padx=4, pady=6)

        # MMM / GGGG / VVV
        ctk.CTkLabel(filters, text="MMM:").pack(side="left", padx=(10, 2), pady=6)
        self.search_mmm_var = ctk.StringVar(value="")
        ctk.CTkEntry(filters, textvariable=self.search_mmm_var, width=70).pack(side="left", padx=4, pady=6)

        ctk.CTkLabel(filters, text="GGGG:").pack(side="left", padx=(10, 2), pady=6)
        self.search_gggg_var = ctk.StringVar(value="")
        ctk.CTkEntry(filters, textvariable=self.search_gggg_var, width=80).pack(side="left", padx=4, pady=6)

        ctk.CTkLabel(filters, text="VVV:").pack(side="left", padx=(10, 2), pady=6)
        self.search_vvv_var = ctk.StringVar(value="")
        ctk.CTkEntry(filters, textvariable=self.search_vvv_var, width=80).pack(side="left", padx=4, pady=6)

        # Mostra OBS (solo nei filtri)
        self.include_obs_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(filters, text="Mostra OBS", variable=self.include_obs_var).pack(side="left", padx=(14, 6), pady=6)

        # Pulsanti filtri
        ctk.CTkButton(filters, text="CERCA", width=120, command=self._search_rc).pack(side="left", padx=(8, 6), pady=6)
        ctk.CTkButton(filters, text="RESET", width=110, command=self._reset_rc_filters).pack(side="left", padx=6, pady=6)

        # --- tabella risultati ---
        columns, headings, props, key_index = self._build_table_schema_with_sw_props()
        self.rc_table = Table(tab, columns=columns, headings=headings, on_double_click=self._doc_dbl, key_index=key_index)
        self.rc_table.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # bind selection -> abilita invio workflow
        self.rc_table.tree.bind("<<TreeviewSelect>>", self._on_rc_select)

        # prima ricerca: elenco base (senza OBS)
        self._refresh_rc_table()

    def _on_rc_select(self, _evt=None):
        code = self._get_selected_rc_code()
        if hasattr(self, "btn_send_to_wf") and self.btn_send_to_wf is not None:
            self.btn_send_to_wf.configure(state="normal" if code else "disabled")

    def _get_selected_rc_code(self) -> str:
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
        code = self._get_selected_rc_code()
        if not code:
            return
        self.wf_code_var.set(code)
        try:
            self.tabs.set("Operativo")
        except Exception:
            pass
        self._refresh_workflow_panel()

    def _reset_rc_filters(self):
        self.search_text_var.set("")
        self.search_state_var.set("")
        self.search_type_var.set("")
        self.search_mmm_var.set("")
        self.search_gggg_var.set("")
        self.search_vvv_var.set("")
        self.include_obs_var.set(False)
        self._refresh_rc_table()

    def _search_rc(self):
        self._refresh_rc_table()

    def _refresh_rc_table(self):
        """Esegue la ricerca/consultazione usando i filtri correnti e aggiorna la tabella unica."""
        # Schema (potrebbe cambiare se modifico le proprieta SW da leggere)
        columns, headings, props, key_index = self._build_table_schema_with_sw_props()
        self.rc_table.set_schema(columns=columns, headings=headings, key_index=key_index)

        txt = (self.search_text_var.get() or "").strip()
        st = (self.search_state_var.get() or "").strip()
        tp = (self.search_type_var.get() or "").strip()
        mmm = (self.search_mmm_var.get() or "").strip().upper()
        gggg = (self.search_gggg_var.get() or "").strip().upper()
        vvv = (self.search_vvv_var.get() or "").strip().upper()
        include_obs = bool(self.include_obs_var.get())

        # query store
        docs = self.store.search_documents(
            text=txt,
            state=st if st else None,
            doc_type=tp if tp else None,
            mmm=mmm if mmm else None,
            gggg=gggg if gggg else None,
            vvv=vvv if vvv else None,
            include_obs=include_obs,
        )

        # bulk sw props values
        sw_values = self.store.get_custom_values_bulk([d.code for d in docs], props) if props else {}

        rows = []
        for d in docs:
            # M: esistenza modello (.sldprt/.sldasm), D: esistenza disegno (.slddrw)
            m_ok, d_ok = self._model_and_drawing_flags(d)

            base = [m_ok, d_ok, d.code, d.doc_type, d.revision, d.state, self._checkout_table_value(d), d.description]
            # append SW dynamic values in same order as props
            extra = []
            for p in props:
                extra.append(sw_values.get(d.code, {}).get(p, ""))
            row_tag = self._state_row_tag(d.state)
            rows.append({"values": (base + extra), "tags": (row_tag,) if row_tag else ()})

        self.rc_table.set_rows(rows)
        self._on_rc_select(None)

    # ESTRATTO IN TabOperativo (pdm_sw/gui/tab_operativo.py)
    # Metodi estratti: _ui_workflow, _build_workflow_panel
    # def _ui_workflow(self): ...
    # (Ora i widget workflow sono creati in TabOperativo, ponti definiti in __init__)

    def _load_selected_doc(self) -> Document | None:
        code = self.wf_code_var.get().strip()
        if not code:
            code = self._get_table_selected_code()
            if code:
                self.wf_code_var.set(code)
        if not code:
            return None
        return self.store.get_document(code)

    def _update_workflow_buttons(self, doc: Document | None) -> None:
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
            _set("wf_btn_checkout", False)
            _set("wf_btn_checkin", False)
            return

        state = str(getattr(doc, "state", "") or "").strip().upper()
        prev_state = str(getattr(doc, "obs_prev_state", "") or "").strip().upper()
        is_checked_out = bool(getattr(doc, "checked_out", False))
        is_checkout_mine = self._is_doc_checked_out_by_me(doc)
        can_checkout_state = state in ("WIP", "IN_REV")

        _set("wf_btn_release", state == "WIP" and is_checked_out and is_checkout_mine)
        _set("wf_btn_create_rev", state == "REL")
        _set("wf_btn_approve", state == "IN_REV" and is_checked_out and is_checkout_mine)
        _set("wf_btn_cancel", state == "IN_REV" and is_checked_out and is_checkout_mine)
        _set("wf_btn_obsolete", state in ("WIP", "REL", "IN_REV"))
        _set("wf_btn_restore_obs", state == "OBS" and prev_state in ("WIP", "REL", "IN_REV"))
        _set("wf_btn_checkout", can_checkout_state and ((not is_checked_out) or is_checkout_mine))
        _set("wf_btn_checkin", can_checkout_state and is_checked_out and is_checkout_mine)

    def _refresh_workflow_panel(self):
        doc = self._load_selected_doc()
        if not doc:
            self.wf_state_var.set("")
            self.wf_info.delete("1.0", "end")
            self._update_workflow_buttons(None)
            return
        self._update_workflow_buttons(doc)
        self.wf_state_var.set(f"Stato: {doc.state}  Rev: {doc.revision:02d}  | {self._checkout_status_label(doc)}")
        self.wf_info.delete("1.0", "end")
        def _shown_path(path_s: str) -> str:
            try:
                return str(path_s or "") if (path_s and Path(path_s).is_file()) else ""
            except Exception:
                return ""
        self.wf_info.insert("end", f"Codice: {doc.code}\n")
        self.wf_info.insert("end", f"Tipo: {doc.doc_type}\n")
        self.wf_info.insert("end", f"MMM/GGGG: {doc.mmm}/{doc.gggg}\n")
        self.wf_info.insert("end", f"Seq: {doc.seq:04d}\n")
        self.wf_info.insert("end", f"VVV: {doc.vvv}\n")
        self.wf_info.insert("end", f"Descrizione: {doc.description}\n")
        created_ts = str(getattr(doc, "created_at", "") or "").replace("T", " ").strip()
        self.wf_info.insert("end", f"Creato il: {created_ts if created_ts else '(n/d)'}\n")
        self.wf_info.insert("end", f"CHECK: {self._checkout_status_label(doc)}\n")
        if getattr(doc, "obs_prev_state", ""):
            self.wf_info.insert("end", f"Stato precedente OBS: {doc.obs_prev_state}\n")
        self.wf_info.insert("end", "\n")
        self.wf_info.insert("end", f"MODEL WIP: {_shown_path(doc.file_wip_path)}\n")
        self.wf_info.insert("end", f"MODEL REL: {_shown_path(doc.file_rel_path)}\n")
        self.wf_info.insert("end", f"MODEL INREV: {_shown_path(doc.file_inrev_path)}\n\n")
        self.wf_info.insert("end", f"DRW WIP: {_shown_path(doc.file_wip_drw_path)}\n")
        self.wf_info.insert("end", f"DRW REL: {_shown_path(doc.file_rel_drw_path)}\n")
        self.wf_info.insert("end", f"DRW INREV: {_shown_path(doc.file_inrev_drw_path)}\n")
        rev_models, rev_drws = self._list_rev_files_for_doc(doc)
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

    def _wf_backup_event(self, reason: str):
        if not self.cfg.backup.enabled:
            return
        res = self.backup.backup_now(reason, force=True)
        if not res.ok:
            warn(res.message)

    def _save_workflow_doc(self, doc: Document) -> None:
        self.store.update_document(
            doc.code,
            state=doc.state,
            revision=doc.revision,
            obs_prev_state=getattr(doc, "obs_prev_state", "") or "",
            file_wip_path=doc.file_wip_path,
            file_rel_path=doc.file_rel_path,
            file_inrev_path=doc.file_inrev_path,
            file_wip_drw_path=doc.file_wip_drw_path,
            file_rel_drw_path=doc.file_rel_drw_path,
            file_inrev_drw_path=doc.file_inrev_drw_path,
        )

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
            text_color="#555555",
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
                warn("Inserisci una nota di almeno 3 caratteri.")
                return
            if len(note) > 2000:
                warn("Nota troppo lunga (massimo 2000 caratteri).")
                return
            result["note"] = note
            top.destroy()

        ctk.CTkButton(btns, text="Annulla", width=120, command=_cancel).pack(side="right", padx=6)
        ctk.CTkButton(btns, text="Conferma", width=120, command=_ok).pack(side="right", padx=6)
        top.bind("<Escape>", lambda _e: _cancel())

        self.wait_window(top)
        return result["note"]

    def _save_workflow_state_note(
        self,
        code: str,
        event_type: str,
        from_state: str,
        to_state: str,
        note: str,
        rev_before: int,
        rev_after: int,
    ) -> None:
        self.store.add_state_note(
            code=code,
            event_type=event_type,
            from_state=from_state,
            to_state=to_state,
            note=note,
            rev_before=rev_before,
            rev_after=rev_after,
        )

    def _close_sw_docs_for_workflow(self, doc: Document) -> None:
        """Best effort: chiude eventuali documenti SW aperti coinvolti nel workflow."""
        try:
            from pdm_sw.sw_api import close_doc, save_existing_doc, _find_running_sw_apps
        except Exception:
            return

        apps = []
        try:
            sw, res = get_solidworks_app(visible=False, timeout_s=3.0, allow_launch=False)
            if res.ok and sw is not None:
                apps.append(sw)
        except Exception:
            pass
        try:
            apps.extend(list(_find_running_sw_apps() or []))
        except Exception:
            pass
        if not apps:
            return

        uniq_apps = {}
        for sw in apps:
            try:
                pid = int(sw.GetProcessID())
            except Exception:
                pid = id(sw)
            uniq_apps[pid] = sw
        apps = list(uniq_apps.values())

        candidates = [
            doc.file_wip_path,
            doc.file_rel_path,
            doc.file_inrev_path,
            doc.file_wip_drw_path,
            doc.file_rel_drw_path,
            doc.file_inrev_drw_path,
        ]
        uniq_paths = []
        seen = set()
        for p in candidates:
            p = str(p or "").strip()
            if not p:
                continue
            key = p.lower()
            if key in seen:
                continue
            seen.add(key)
            uniq_paths.append(p)

        for sw in apps:
            for p in uniq_paths:
                od = None
                try:
                    if hasattr(sw, "GetOpenDocumentByName"):
                        od = sw.GetOpenDocumentByName(p)
                except Exception:
                    od = None
                if od is None:
                    try:
                        if hasattr(sw, "GetOpenDocumentByName"):
                            od = sw.GetOpenDocumentByName(Path(p).name)
                    except Exception:
                        od = None
                if od is not None:
                    try:
                        save_existing_doc(od)
                    except Exception:
                        pass
                try:
                    close_doc(sw, doc=od, file_path=p)
                except Exception:
                    pass

    def _run_workflow_transition(self, action_label: str, fn, *args, **kwargs):
        locked_code = ""
        has_lock = False
        if args and isinstance(args[0], Document):
            locked_code = str(args[0].code or "").strip()
            ok, _holder = self._acquire_doc_lock(locked_code, action=f"WF_{action_label.upper()}")
            if not ok:
                return None
            has_lock = True
            try:
                self._close_sw_docs_for_workflow(args[0])
            except Exception:
                pass
            try:
                kwargs.setdefault("log_file", str(self._workflow_log_path()))
            except Exception:
                pass
        try:
            out = fn(*args, **kwargs)
            if locked_code:
                self._log_activity(action=f"WF_{action_label.upper()}", code=locked_code, status="OK", message="Transizione eseguita.")
            return out
        except FileExistsError as e:
            if locked_code:
                self._log_activity(action=f"WF_{action_label.upper()}", code=locked_code, status="ERROR", message=f"File exists: {e}")
            warn(
                "Operazione non riuscita: file destinazione gia presente.\n"
                "Verifica storico revisioni/cartelle e riprova.\n\n"
                f"Dettaglio: {e}"
            )
            return None
        except PermissionError as e:
            if locked_code:
                self._log_activity(action=f"WF_{action_label.upper()}", code=locked_code, status="ERROR", message=f"Permission: {e}")
            warn(
                "Operazione non riuscita: uno o piu file sono in uso o non accessibili.\n"
                "Chiudi i file in SolidWorks/Explorer e verifica i permessi del file/cartella.\n\n"
                f"Dettaglio: {e}"
            )
            return None
        except Exception as e:
            if locked_code:
                self._log_activity(action=f"WF_{action_label.upper()}", code=locked_code, status="ERROR", message=str(e))
            warn(f"Errore durante {action_label}: {e}")
            return None
        finally:
            if has_lock and locked_code:
                self._release_doc_lock(locked_code)

    def _wf_release(self):
        doc = self._load_selected_doc()
        if not doc:
            return
        if str(doc.state or "").strip().upper() != "WIP":
            warn("Per release serve stato WIP.")
            return
        if not self._require_wip_checkout(doc, "Release"):
            return
        from_state = doc.state
        rev_before = int(doc.revision)
        note = self._prompt_workflow_note(doc.code, "Release", from_state, "REL")
        if note is None:
            return
        out = self._run_workflow_transition("release", release_wip, doc, self.cfg.solidworks.archive_root)
        if out is None:
            return
        doc2, res = out
        if not res.ok:
            warn(res.message)
            return
        self._save_workflow_doc(doc2)
        try:
            self._save_workflow_state_note(
                code=doc.code,
                event_type="RELEASE",
                from_state=from_state,
                to_state=doc2.state,
                note=note,
                rev_before=rev_before,
                rev_after=int(doc2.revision),
            )
        except Exception as e:
            warn(f"Cambio stato eseguito, ma salvataggio nota fallito: {e}")
        try:
            self.store.clear_document_checkout(doc.code)
        except Exception:
            pass
        self._wf_backup_event("release")
        self._refresh_all()

    def _wf_create_rev(self):
        doc = self._load_selected_doc()
        if not doc:
            return
        if doc.state != "REL":
            warn("Per creare revisione serve stato REL.")
            return
        from_state = doc.state
        rev_before = int(doc.revision)
        note = self._prompt_workflow_note(doc.code, "Crea revisione", from_state, "IN_REV")
        if note is None:
            return
        out = self._run_workflow_transition("creazione revisione", create_inrev, doc, self.cfg.solidworks.archive_root)
        if out is None:
            return
        doc2, res = out
        if not res.ok:
            warn(res.message)
            return
        self._save_workflow_doc(doc2)
        try:
            self._save_workflow_state_note(
                code=doc.code,
                event_type="CREATE_REV",
                from_state=from_state,
                to_state=doc2.state,
                note=note,
                rev_before=rev_before,
                rev_after=int(doc2.revision),
            )
        except Exception as e:
            warn(f"Cambio stato eseguito, ma salvataggio nota fallito: {e}")
        self._wf_backup_event("create_rev")
        self._refresh_all()

    def _wf_approve(self):
        doc = self._load_selected_doc()
        if not doc:
            return
        if doc.state != "IN_REV":
            warn("Per approvare serve stato IN_REV.")
            return
        if not self._require_checkout_for_edit(doc, "Approvazione revisione"):
            return
        from_state = doc.state
        rev_before = int(doc.revision)
        note = self._prompt_workflow_note(doc.code, "Approva revisione", from_state, "REL")
        if note is None:
            return
        out = self._run_workflow_transition("approvazione revisione", approve_inrev, doc, self.cfg.solidworks.archive_root)
        if out is None:
            return
        doc2, res = out
        if not res.ok:
            warn(res.message)
            return
        self._save_workflow_doc(doc2)
        try:
            self._save_workflow_state_note(
                code=doc.code,
                event_type="APPROVE_REV",
                from_state=from_state,
                to_state=doc2.state,
                note=note,
                rev_before=rev_before,
                rev_after=int(doc2.revision),
            )
        except Exception as e:
            warn(f"Cambio stato eseguito, ma salvataggio nota fallito: {e}")
        try:
            self.store.clear_document_checkout(doc.code)
        except Exception:
            pass
        self._wf_backup_event("approve_rev")
        self._refresh_all()

    def _wf_cancel(self):
        doc = self._load_selected_doc()
        if not doc:
            return
        if doc.state != "IN_REV":
            warn("Per annullare serve stato IN_REV.")
            return
        if not self._require_checkout_for_edit(doc, "Annullamento revisione"):
            return
        from_state = doc.state
        rev_before = int(doc.revision)
        note = self._prompt_workflow_note(doc.code, "Annulla revisione", from_state, "REL")
        if note is None:
            return
        out = self._run_workflow_transition("annullamento revisione", cancel_inrev, doc)
        if out is None:
            return
        doc2, res = out
        if not res.ok:
            warn(res.message)
            return
        self._save_workflow_doc(doc2)
        try:
            self._save_workflow_state_note(
                code=doc.code,
                event_type="CANCEL_REV",
                from_state=from_state,
                to_state=doc2.state,
                note=note,
                rev_before=rev_before,
                rev_after=int(doc2.revision),
            )
        except Exception as e:
            warn(f"Cambio stato eseguito, ma salvataggio nota fallito: {e}")
        try:
            self.store.clear_document_checkout(doc.code)
        except Exception:
            pass
        self._wf_backup_event("cancel_rev")
        self._refresh_all()

    def _wf_obsolete(self):
        doc = self._load_selected_doc()
        if not doc:
            return
        prev_state = doc.state
        rev_before = int(doc.revision)
        note = self._prompt_workflow_note(doc.code, "Imposta OBS", prev_state, "OBS")
        if note is None:
            return
        out = self._run_workflow_transition("impostazione OBS", set_obsolete, doc)
        if out is None:
            return
        doc2, res = out
        if not res.ok:
            warn(res.message)
            return
        doc2.obs_prev_state = prev_state if prev_state in ("WIP", "REL", "IN_REV") else ""
        self._save_workflow_doc(doc2)
        try:
            self._save_workflow_state_note(
                code=doc.code,
                event_type="SET_OBSOLETE",
                from_state=prev_state,
                to_state=doc2.state,
                note=note,
                rev_before=rev_before,
                rev_after=int(doc2.revision),
            )
        except Exception as e:
            warn(f"Cambio stato eseguito, ma salvataggio nota fallito: {e}")
        try:
            self.store.clear_document_checkout(doc.code)
        except Exception:
            pass
        self._wf_backup_event("obsolete")
        self._refresh_all()

    def _wf_restore_obs(self):
        doc = self._load_selected_doc()
        if not doc:
            return
        if doc.state != "OBS":
            warn("Il documento non e in stato OBS.")
            return
        prev_state = (getattr(doc, "obs_prev_state", "") or "").strip().upper()
        if not prev_state:
            warn("Stato precedente OBS non disponibile.")
            return
        rev_before = int(doc.revision)
        note = self._prompt_workflow_note(doc.code, "Ripristina da OBS", "OBS", prev_state)
        if note is None:
            return
        out = self._run_workflow_transition("ripristino da OBS", restore_obsolete, doc, prev_state)
        if out is None:
            return
        doc2, res = out
        if not res.ok:
            warn(res.message)
            return
        doc2.obs_prev_state = ""
        self._save_workflow_doc(doc2)
        try:
            self._save_workflow_state_note(
                code=doc.code,
                event_type="RESTORE_OBS",
                from_state="OBS",
                to_state=doc2.state,
                note=note,
                rev_before=rev_before,
                rev_after=int(doc2.revision),
            )
        except Exception as e:
            warn(f"Cambio stato eseguito, ma salvataggio nota fallito: {e}")
        try:
            self.store.clear_document_checkout(doc.code)
        except Exception:
            pass
        self._wf_backup_event("restore_obs")
        self._refresh_all()

    # ---------------- Tab: Monitor
    # ESTRATTO IN: pdm_sw/gui/tab_monitor.py -> TabMonitor
    # def _ui_monitor(self):
    #     tab = self.tab_monitor
    #
    #     self.monitor_auto_var = tk.BooleanVar(value=True)
    #     self.monitor_limit_var = tk.StringVar(value="200")
    #     self.monitor_after_id = None
    #
    #     actions = ctk.CTkFrame(tab, fg_color="#EAF2FF")
    #     actions.pack(fill="x", padx=10, pady=(10, 6))
    #
    #     ctk.CTkButton(actions, text="AGGIORNA ORA", width=140, command=self._refresh_monitor_panel).pack(side="left", padx=6, pady=6)
    #     ctk.CTkCheckBox(actions, text="Auto refresh 5s", variable=self.monitor_auto_var, command=self._refresh_monitor_panel).pack(side="left", padx=12, pady=6)
    #     ctk.CTkLabel(actions, text="Attivita recenti:").pack(side="left", padx=(12, 4), pady=6)
    #     ctk.CTkComboBox(actions, variable=self.monitor_limit_var, values=["100", "200", "500", "1000"], width=90, command=lambda _=None: self._refresh_monitor_panel()).pack(side="left", padx=4, pady=6)
    #
    #     self.monitor_summary_var = tk.StringVar(value="Lock attivi: 0 | Eventi: 0")
    #     ctk.CTkLabel(actions, textvariable=self.monitor_summary_var, font=ctk.CTkFont(size=12, weight="bold")).pack(side="right", padx=10, pady=6)
    #
    #     locks_box = ctk.CTkFrame(tab)
    #     locks_box.pack(fill="both", expand=True, padx=10, pady=(0, 6))
    #     ctk.CTkLabel(locks_box, text="Lock Attivi", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=8, pady=(8, 4))
    #     lock_cols = ["code", "owner_user", "owner_host", "acquired_at", "updated_at", "expires_at", "remaining_min"]
    #     lock_heads = ["CODICE", "UTENTE", "HOST", "ACQUISITO", "AGGIORNATO", "SCADENZA", "MIN RESIDUI"]
    #     self.monitor_lock_table = Table(locks_box, columns=lock_cols, headings=lock_heads, key_index=0)
    #     self.monitor_lock_table.pack(fill="both", expand=True, padx=6, pady=(0, 8))
    #     self.monitor_lock_table.tree.tag_configure("lock_mine", foreground="#0B5ED7")
    #
    #     activity_box = ctk.CTkFrame(tab)
    #     activity_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    #     ctk.CTkLabel(activity_box, text="Activity Log", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=8, pady=(8, 4))
    #     act_cols = ["created_at", "action", "status", "code", "user_display", "host", "file_path", "message"]
    #     act_heads = ["DATA/ORA", "AZIONE", "ESITO", "CODICE", "UTENTE", "HOST", "FILE", "MESSAGGIO"]
    #     self.monitor_activity_table = Table(activity_box, columns=act_cols, headings=act_heads, key_index=3)
    #     self.monitor_activity_table.pack(fill="both", expand=True, padx=6, pady=(0, 8))
    #     try:
    #         self.monitor_activity_table.tree.column("file_path", width=360, stretch=True)
    #         self.monitor_activity_table.tree.column("message", width=260, stretch=True)
    #     except Exception:
    #         pass
    #     self.monitor_activity_table.tree.tag_configure("act_error", foreground="#B91C1C")
    #     self.monitor_activity_table.tree.tag_configure("act_warn", foreground="#B8860B")
    #     self.monitor_activity_table.tree.tag_configure("act_ok", foreground="#166534")
    #
    # def _refresh_monitor_panel(self):
    #     if not hasattr(self, "monitor_lock_table") or not hasattr(self, "monitor_activity_table"):
    #         return
    #
    #     try:
    #         limit = int((self.monitor_limit_var.get() or "200").strip())
    #     except Exception:
    #         limit = 200
    #     limit = min(1000, max(50, limit))
    #
    #     my_session = str(self.session.get("session_id", ""))
    #     now = datetime.now()
    #
    #     try:
    #         locks = self.store.list_active_locks(limit=500)
    #     except Exception:
    #         locks = []
    #     lock_rows = []
    #     for lk in locks:
    #         exp = str(lk.get("expires_at", ""))
    #         rem = ""
    #         try:
    #             rem_min = max(0.0, (datetime.fromisoformat(exp) - now).total_seconds() / 60.0)
    #             rem = f"{rem_min:.1f}"
    #         except Exception:
    #             rem = ""
    #         vals = [
    #             str(lk.get("code", "")),
    #             str(lk.get("owner_user", "")),
    #             str(lk.get("owner_host", "")),
    #             str(lk.get("acquired_at", "")).replace("T", " "),
    #             str(lk.get("updated_at", "")).replace("T", " "),
    #             exp.replace("T", " "),
    #             rem,
    #         ]
    #         tags = ("lock_mine",) if str(lk.get("owner_session", "")) == my_session else ()
    #         lock_rows.append({"values": vals, "tags": tags})
    #     self.monitor_lock_table.set_rows(lock_rows)
    #
    #     try:
    #         activities = self.store.list_recent_activity(limit=limit)
    #     except Exception:
    #         activities = []
    #     act_rows = []
    #     for a in activities:
    #         status = str(a.get("status", "")).strip().upper()
    #         details = a.get("details", {}) if isinstance(a.get("details", {}), dict) else {}
    #         file_path = str(details.get("path", "") or "")
    #         tag = ""
    #         if status in ("ERROR", "KO"):
    #             tag = "act_error"
    #         elif status in ("WARN", "WARNING", "LOCKED"):
    #             tag = "act_warn"
    #         elif status == "OK":
    #             tag = "act_ok"
    #         vals = [
    #             str(a.get("created_at", "")).replace("T", " "),
    #             str(a.get("action", "")),
    #             status,
    #             str(a.get("code", "")),
    #             str(a.get("user_display", "")),
    #             str(a.get("host", "")),
    #             file_path,
    #             str(a.get("message", "")),
    #         ]
    #         act_rows.append({"values": vals, "tags": ((tag,) if tag else ())})
    #     self.monitor_activity_table.set_rows(act_rows)
    #
    #     self.monitor_summary_var.set(f"Lock attivi: {len(lock_rows)} | Eventi: {len(act_rows)}")
    #
    #     if getattr(self, "monitor_after_id", None):
    #         try:
    #             self.after_cancel(self.monitor_after_id)
    #         except Exception:
    #             pass
    #         self.monitor_after_id = None
    #     if bool(self.monitor_auto_var.get()):
    #         self.monitor_after_id = self.after(5000, self._refresh_monitor_panel)


    # ===============================================================================
    # METODO DEPRECATO - Estratto in pdm_sw.gui.tab_manuale.TabManuale
    # Mantenuto temporaneamente per riferimento durante il refactoring
    # ===============================================================================
    # def _ui_manuale(self):
    #     tab = self.tab_manuale
    #     try:
    #         tab.configure(fg_color="#FFF4B8")
    #     except Exception:
    #         pass
    #
    #     ctk.CTkLabel(
    #         tab,
    #         text=f"Manuale Rapido (Rev {APP_REV})",
    #         font=ctk.CTkFont(size=16, weight="bold"),
    #     ).pack(anchor="w", padx=12, pady=(12, 8))
    #
    #     ctk.CTkLabel(
    #         tab,
    #         text="I dettagli completi sono nel file README.md.",
    #         text_color="#555555",
    #     ).pack(anchor="w", padx=12, pady=(0, 8))
    #
    #     box = ctk.CTkTextbox(tab)
    #     box.pack(fill="both", expand=True, padx=12, pady=(0, 12))
    #     manual_text = (
    #         "PANORAMICA FUNZIONALITA\n"
    #         "=======================\n"
    #         "1) Workspace\n"
    #         "- Usa il pulsante WORKSPACE... in alto per aprire gli strumenti workspace.\n"
    #         "- Da li puoi cambiare/creare/copiare/cancellare workspace e scegliere la cartella condivisa.\n"
    #         "- Ogni workspace ha configurazione, database e backup separati.\n"
    #         "- Puoi scegliere la CARTELLA CONDIVISA (default: cartella attuale).\n"
    #         "- La root condivisa attiva e mostrata in alto come SHARED: <percorso>.\n"
    #         "- Le funzioni di configurazione sono raccolte nella tab principale SETUP.\n"
    #         "\n"
    #         "2) Gestione codifica\n"
    #         "- Configura formato codice: MMM, GGGG, VVV e progressivo.\n"
    #         "- Definisci separatori, lunghezze e regole di validazione.\n"
    #         "\n"
    #         "3) Generatore codici\n"
    #         "- Gestisci elenco macchine (MMM) e gruppi (GGGG).\n"
    #         "\n"
    #         "4) SolidWorks\n"
    #         "- Imposta archivio e template PART/ASSY/DRW.\n"
    #         "- Configura mappatura proprieta PDM -> SW e lettura SW -> PDM.\n"
    #         "- Pubblica la macro bootstrap per la workspace corrente.\n"
    #         "\n"
    #         "5) Codifica\n"
    #         "- Crea solo codice (WIP) oppure codice + file da template.\n"
    #         "- Importa file esistenti (.sldprt/.sldasm/.slddrw) nel WIP.\n"
    #         "\n"
    #         "6) Gerarchia\n"
    #         "- Vista ad albero MMM -> GGGG -> CODICE.\n"
    #         "\n"
    #         "7) Operativo\n"
    #         "- Tab unica per Ricerca&Consultazione + Workflow.\n"
    #         "- Sinistra: filtri, tabella risultati, apertura MODEL/DRW, creazione file mancanti e comandi FORZA PDM->SW / SW->PDM.\n"
    #         "- Destra: pannello workflow con transizioni stato e dettagli del codice selezionato.\n"
    #         "- Il pulsante INVIA A WORKFLOW e nel pannello workflow (in alto).\n"
    #         "- Puoi regolare la larghezza del pannello workflow trascinando il separatore centrale (salvata in locale).\n"
    #         "- Transizioni stato: WIP -> REL -> IN_REV -> REL e OBS.\n"
    #         "- Ripristino da OBS allo stato precedente se disponibile.\n"
    #         "- Ogni cambio stato richiede una nota obbligatoria (minimo 3 caratteri).\n"
    #         "- Le note sono salvate con data/ora, tipo evento e stati (da -> a).\n"
    #         "- Ogni transizione registra operazioni file in WORKSPACES\\<workspace>\\LOGS\\workflow.log.\n"
    #         "\n"
    #         "8) Monitor\n"
    #         "- Vista LOCK ATTIVI: codice, utente, host, scadenza lock e minuti residui.\n"
    #         "- Vista ACTIVITY LOG: data/ora, azione, esito, codice, utente, host, file e messaggio.\n"
    #         "- Auto refresh ogni 5 secondi (disattivabile) e filtro numero eventi recenti.\n"
    #         "\n"
    #         "9) Multiutente (3-6 utenti)\n"
    #         "- I lock sono condivisi nel DB workspace e proteggono le transizioni workflow.\n"
    #         "- Se un utente ha lock su un codice, gli altri vedono il blocco e non possono eseguire la stessa transizione.\n"
    #         "- I lock vengono rilasciati a fine operazione o alla chiusura sessione; in caso crash scadono a timeout.\n"
    #         "- Tracciamento aperture file (livello 1): log da tab Ricerca&Consultazione (doppio click/APRI MODELLO/APRI DRW).\n"
    #         "\n"
    #         "INSTALLAZIONE MACRO SOLIDWORKS (COMPLETA)\n"
    #         "=========================================\n"
    #         "Prerequisiti:\n"
    #         "- SolidWorks installato.\n"
    #         "- Tab SolidWorks compilata (Archivio + template) e salvata.\n"
    #         "\n"
    #         "Passi:\n"
    #         "1) Vai in tab SolidWorks e premi: PUBBLICA MACRO SOLIDWORKS.\n"
    #         "2) Apri il file istruzioni generato in:\n"
    #         "   SW_MACROS\\INSTALL_MACRO_<workspace_id>.txt\n"
    #         "3) In SolidWorks: Strumenti > Macro > Nuova...\n"
    #         "4) Salva la macro in:\n"
    #         "   SW_MACROS\\PDM_SW_BOOTSTRAP_<workspace_id>.swp\n"
    #         "5) Nell'editor VBA: File > Import File... e importa:\n"
    #         "   SW_MACROS\\PDM_SW_BOOTSTRAP_<workspace_id>.bas\n"
    #         "6) Salva e chiudi VBA.\n"
    #         "7) (Consigliato) Aggiungi pulsante toolbar:\n"
    #         "   Strumenti > Personalizza > Comandi > Macro > Esegui Macro.\n"
    #         "\n"
    #         "Build EXE payload (opzionale ma consigliato):\n"
    #         "1) Apri:\n"
    #         "   WORKSPACES\\<workspace_folder>\\macros\\payload\n"
    #         "2) Esegui:\n"
    #         "   build_payload_exe.bat\n"
    #         "3) Verifica presenza di PDM_SW_PAYLOAD.exe nella cartella payload.\n"
    #         "\n"
    #         "Diagnostica macro:\n"
    #         "- Log bootstrap:\n"
    #         "  SW_CACHE\\<workspace_id>\\payload\\bootstrap.log\n"
    #         "- Log payload:\n"
    #         "  SW_CACHE\\<workspace_id>\\payload\\payload.log\n"
    #     )
    #     box.insert(
    #         "1.0",
    #         manual_text,
    #     )
    #     box.configure(state="disabled")


    def _call_safe(self, method_name: str):
        fn = getattr(self, method_name, None)
        if not callable(fn):
            warn(f"Funzione non disponibile: {method_name}")
            return
        try:
            fn()
        except Exception as e:
            warn(f"Errore in {method_name}: {e}")

    def _workspace_tools_dialog(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Workspace e Cartella Condivisa")
        dlg.geometry("560x380")
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="Strumenti Workspace", font=ctk.CTkFont(size=15, weight="bold")).pack(anchor="w", padx=12, pady=(12, 6))
        ctk.CTkLabel(dlg, text=f"Workspace attiva: {self.ws.name}").pack(anchor="w", padx=12, pady=(0, 2))
        ctk.CTkLabel(dlg, text=f"Shared root: {self.shared_root}", text_color="#555555").pack(anchor="w", padx=12, pady=(0, 10))

        grid = ctk.CTkFrame(dlg, fg_color="transparent")
        grid.pack(fill="x", padx=12, pady=6)
        grid.grid_columnconfigure(0, weight=1)
        grid.grid_columnconfigure(1, weight=1)

        def _open_and_close(method_name: str):
            dlg.destroy()
            self._call_safe(method_name)

        ctk.CTkButton(grid, text="CAMBIA WORKSPACE", command=lambda: _open_and_close("_change_workspace_dialog")).grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        ctk.CTkButton(grid, text="CARTELLA CONDIVISA", command=lambda: _open_and_close("_change_shared_root_dialog")).grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        ctk.CTkButton(grid, text="CREA WORKSPACE", command=lambda: _open_and_close("_create_workspace_dialog")).grid(row=1, column=0, sticky="ew", padx=6, pady=6)
        ctk.CTkButton(grid, text="COPIA WORKSPACE", command=lambda: _open_and_close("_copy_workspace_dialog")).grid(row=1, column=1, sticky="ew", padx=6, pady=6)
        ctk.CTkButton(
            grid,
            text="CANCELLA WORKSPACE",
            fg_color="#b91c1c",
            hover_color="#991b1b",
            command=lambda: _open_and_close("_delete_workspace_dialog"),
        ).grid(row=2, column=0, columnspan=2, sticky="ew", padx=6, pady=6)
        ctk.CTkButton(
            grid,
            text="MIGRA ARCHIVIO CAD",
            fg_color="#0B5ED7",
            hover_color="#0A58CA",
            command=lambda: _open_and_close("_migrate_archive_layout_dialog"),
        ).grid(row=3, column=0, columnspan=2, sticky="ew", padx=6, pady=6)

        ctk.CTkButton(dlg, text="Chiudi", width=120, command=dlg.destroy).pack(side="right", padx=12, pady=12)

    def _migrate_archive_layout_dialog(self):
        archive_root = str(getattr(self.cfg.solidworks, "archive_root", "") or "").strip()
        if not archive_root:
            warn("Archivio non configurato (tab SolidWorks).")
            return
        self._log_activity("ARCHIVE_MIGRATION_DRYRUN", status="OK", message="Avvio analisi layout archivio.")

        try:
            preview = run_archive_layout_migration(
                store=self.store,
                archive_root=archive_root,
                apply_changes=False,
            )
        except Exception as e:
            warn(f"Analisi migrazione fallita: {e}")
            return

        if not preview.get("ok", False):
            errs = preview.get("errors", []) or []
            msg = "Analisi migrazione completata con errori.\n"
            if errs:
                msg += "\n" + "\n".join(str(x) for x in errs[:12])
            warn(msg)
            return

        conflicts = preview.get("conflicts", []) or []
        sample_moves = preview.get("sample_moves", []) or []
        msg = (
            "Analisi layout archivio completata.\n\n"
            f"Documenti analizzati: {int(preview.get('docs_scanned', 0))}\n"
            f"Documenti da aggiornare DB: {int(preview.get('docs_to_update', 0))}\n"
            f"Spostamenti pianificati: {int(preview.get('moves_planned', 0))}\n"
            f"Conflitti rilevati: {len(conflicts)}\n\n"
        )
        if sample_moves:
            msg += "Esempi spostamenti:\n" + "\n".join(str(x) for x in sample_moves[:8]) + "\n\n"
        if conflicts:
            msg += "Conflitti (estratto):\n" + "\n".join(str(x) for x in conflicts[:8]) + "\n\n"
            msg += "Risolvi i conflitti e riesegui.\n"
            warn(msg)
            return

        if not ask(msg + "Eseguire ORA la migrazione archivio?"):
            return

        bkp = self.backup.backup_now("archive_layout_migration", force=True)
        if not bkp.ok:
            self._log_activity("ARCHIVE_MIGRATION", status="ERROR", message=f"Backup fallito: {bkp.message}")
            warn(f"Backup pre-migrazione fallito: {bkp.message}")
            return

        try:
            result = run_archive_layout_migration(
                store=self.store,
                archive_root=archive_root,
                apply_changes=True,
            )
        except Exception as e:
            self._log_activity("ARCHIVE_MIGRATION", status="ERROR", message=str(e))
            warn(f"Migrazione archivio fallita: {e}")
            return

        if not result.get("ok", False):
            errs = result.get("errors", []) or []
            self._log_activity("ARCHIVE_MIGRATION", status="ERROR", message="Migrazione con errori.")
            warn(
                "Migrazione completata con errori.\n\n"
                + "\n".join(str(x) for x in errs[:12])
            )
            self._refresh_all()
            return

        self._log_activity(
            "ARCHIVE_MIGRATION",
            status="OK",
            message=f"done={int(result.get('moves_done', 0))} docs={int(result.get('docs_updated', 0))}",
        )
        self._refresh_all()
        info(
            "Migrazione archivio completata.\n\n"
            f"Documenti aggiornati: {int(result.get('docs_updated', 0))}\n"
            f"Spostamenti eseguiti: {int(result.get('moves_done', 0))}\n"
            f"Spostamenti mancanti (sorgente assente): {int(result.get('moves_missing', 0))}\n"
            f"Backup: {bkp.path or '(ok)'}"
        )

    def _change_workspace_dialog(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Cambia workspace")
        dlg.geometry("520x190")
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="Seleziona una WORKSPACE", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=12, pady=(12, 8))

        ws_list = self.ws_mgr.list()
        ws_names = [f"{w.name} - {w.description}" for w in ws_list]
        ws_ids = [w.id for w in ws_list]

        cur_label = ""
        for i, wid in enumerate(ws_ids):
            if wid == self.ws_id:
                cur_label = ws_names[i]
                break
        pick_var = tk.StringVar(value=cur_label if cur_label else (ws_names[0] if ws_names else ""))

        f0 = ctk.CTkFrame(dlg)
        f0.pack(fill="x", padx=12, pady=6)
        ctk.CTkLabel(f0, text="Workspace", width=100, anchor="w").pack(side="left", padx=6)
        ctk.CTkOptionMenu(f0, values=ws_names or [""], variable=pick_var).pack(side="left", fill="x", expand=True, padx=6)

        def _ok():
            if not ws_names:
                return
            try:
                idx = ws_names.index(pick_var.get())
            except Exception:
                idx = 0
            ws_id = ws_ids[idx]
            dlg.destroy()
            self._switch_workspace(ws_id)

        ctk.CTkButton(dlg, text="Attiva", command=_ok, width=120).pack(side="right", padx=12, pady=12)

    def _create_workspace_dialog(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Crea workspace")
        dlg.geometry("560x280")
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="Crea una nuova WORKSPACE", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=12, pady=(12, 8))

        name_var = tk.StringVar(value="")
        desc_var = tk.StringVar(value="")
        copy_var = tk.BooleanVar(value=True)

        f0 = ctk.CTkFrame(dlg)
        f0.pack(fill="x", padx=12, pady=6)
        ctk.CTkLabel(f0, text="Nome", width=100, anchor="w").pack(side="left", padx=6)
        ctk.CTkEntry(f0, textvariable=name_var).pack(side="left", fill="x", expand=True, padx=6)

        f1 = ctk.CTkFrame(dlg)
        f1.pack(fill="x", padx=12, pady=6)
        ctk.CTkLabel(f1, text="Descrizione", width=100, anchor="w").pack(side="left", padx=6)
        ctk.CTkEntry(f1, textvariable=desc_var).pack(side="left", fill="x", expand=True, padx=6)

        ctk.CTkCheckBox(dlg, text="Copia configurazione e database dalla workspace corrente", variable=copy_var).pack(anchor="w", padx=16, pady=8)

        def _ok():
            name = (name_var.get() or "").strip()
            desc = (desc_var.get() or "").strip()
            if not name:
                warn("Inserisci il nome workspace.")
                return
            try:
                if copy_var.get():
                    ws = self.ws_mgr.copy(self.ws_id, name, desc, copy_db=True)
                else:
                    ws = self.ws_mgr.create(name, desc)
            except Exception as e:
                warn(f"Creazione workspace fallita: {e}")
                return
            dlg.destroy()
            self._switch_workspace(ws.id)

        ctk.CTkButton(dlg, text="Crea", command=_ok, width=120).pack(side="right", padx=12, pady=12)

    def _copy_workspace_dialog(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Copia workspace")
        dlg.geometry("600x320")
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="Copia una WORKSPACE", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=12, pady=(12, 8))

        ws_list = self.ws_mgr.list()
        ws_names = [f"{w.name} - {w.description}" for w in ws_list]
        ws_ids = [w.id for w in ws_list]

        src_var = tk.StringVar(value=ws_names[0] if ws_names else "")
        name_var = tk.StringVar(value="")
        desc_var = tk.StringVar(value="")
        copy_db_var = tk.BooleanVar(value=True)

        f0 = ctk.CTkFrame(dlg)
        f0.pack(fill="x", padx=12, pady=6)
        ctk.CTkLabel(f0, text="Sorgente", width=100, anchor="w").pack(side="left", padx=6)
        ctk.CTkOptionMenu(f0, values=ws_names or [""], variable=src_var).pack(side="left", fill="x", expand=True, padx=6)

        f1 = ctk.CTkFrame(dlg)
        f1.pack(fill="x", padx=12, pady=6)
        ctk.CTkLabel(f1, text="Nuovo nome", width=100, anchor="w").pack(side="left", padx=6)
        ctk.CTkEntry(f1, textvariable=name_var).pack(side="left", fill="x", expand=True, padx=6)

        f2 = ctk.CTkFrame(dlg)
        f2.pack(fill="x", padx=12, pady=6)
        ctk.CTkLabel(f2, text="Descrizione", width=100, anchor="w").pack(side="left", padx=6)
        ctk.CTkEntry(f2, textvariable=desc_var).pack(side="left", fill="x", expand=True, padx=6)

        ctk.CTkCheckBox(dlg, text="Copia anche database", variable=copy_db_var).pack(anchor="w", padx=16, pady=8)

        def _ok():
            if not ws_names:
                return
            try:
                idx = ws_names.index(src_var.get())
            except Exception:
                idx = 0
            src_id = ws_ids[idx]
            name = (name_var.get() or "").strip()
            desc = (desc_var.get() or "").strip()
            if not name:
                warn("Inserisci il nome della nuova workspace.")
                return
            try:
                ws = self.ws_mgr.copy(src_id, name, desc, copy_db=bool(copy_db_var.get()))
            except Exception as e:
                warn(f"Copia workspace fallita: {e}")
                return
            dlg.destroy()
            self._switch_workspace(ws.id)

        ctk.CTkButton(dlg, text="Copia", command=_ok, width=120).pack(side="right", padx=12, pady=12)

    def _delete_workspace_dialog(self):
        dlg = ctk.CTkToplevel(self)
        dlg.title("Cancella workspace")
        dlg.geometry("620x320")
        dlg.grab_set()

        ctk.CTkLabel(dlg, text="Cancella una WORKSPACE", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=12, pady=(12, 8))

        ws_list = self.ws_mgr.list()
        ws_names = [f"{w.name} - {w.description}" for w in ws_list]
        ws_ids = [w.id for w in ws_list]

        pick_var = tk.StringVar(value=ws_names[0] if ws_names else "")
        del_folder_var = tk.BooleanVar(value=False)
        confirm_var = tk.StringVar(value="")

        f0 = ctk.CTkFrame(dlg)
        f0.pack(fill="x", padx=12, pady=6)
        ctk.CTkLabel(f0, text="Workspace", width=100, anchor="w").pack(side="left", padx=6)
        ctk.CTkOptionMenu(f0, values=ws_names or [""], variable=pick_var).pack(side="left", fill="x", expand=True, padx=6)

        ctk.CTkCheckBox(dlg, text="Elimina anche la cartella su disco", variable=del_folder_var).pack(anchor="w", padx=16, pady=6)

        ctk.CTkLabel(dlg, text="Scrivi 'ELIMINA' per confermare:").pack(anchor="w", padx=16, pady=(10, 4))
        ctk.CTkEntry(dlg, textvariable=confirm_var).pack(fill="x", padx=16)

        def _ok():
            if not ws_names:
                return
            if confirm_var.get().strip().upper() != "ELIMINA":
                warn("Conferma non valida.")
                return
            try:
                idx = ws_names.index(pick_var.get())
            except Exception:
                idx = 0
            ws_id = ws_ids[idx]
            if ws_id == self.ws_id and len(ws_ids) == 1:
                warn("Impossibile eliminare l'ultima workspace.")
                return
            if not ask("Confermi eliminazione?"):
                return
            self.ws_mgr.delete(ws_id, delete_folder=bool(del_folder_var.get()))
            dlg.destroy()
            self._switch_workspace(self.ws_mgr.ensure_default().id)

        ctk.CTkButton(dlg, text="Cancella", fg_color="#b91c1c", hover_color="#991b1b", command=_ok).pack(pady=14)

    def _rebind_workspace_context_to_tabs(self, refresh_sw_tab: bool = False) -> None:
        """Aggiorna i riferimenti cfg/store nei tab in base alla workspace corrente."""
        tab_names = (
            "tab_generatore_obj",
            "tab_codifica_obj",
            "tab_gestione_codifica_obj",
            "tab_solidworks_obj",
            "tab_gerarchia_obj",
            "tab_monitor_obj",
            "tab_operativo_obj",
            "tab_manuale_obj",
        )
        for tab_name in tab_names:
            tab = getattr(self, tab_name, None)
            if not tab:
                continue
            try:
                tab.store = self.store
                tab.cfg = self.cfg
            except Exception:
                pass
        if refresh_sw_tab and getattr(self, "tab_solidworks_obj", None):
            try:
                self.tab_solidworks_obj.refresh()
            except Exception:
                pass

    # ---------------- Workspace switch
    def _switch_workspace(self, ws_id: str):
        if ws_id == self.ws_id:
            return
        old_ws = self.ws_id
        if getattr(self, "monitor_after_id", None):
            try:
                self.after_cancel(self.monitor_after_id)
            except Exception:
                pass
            self.monitor_after_id = None

        # backup on switch (se dirty) + daily if enabled
        if self.cfg.backup.enabled:
            if self.cfg.backup.daily_enabled:
                self.backup.maybe_daily_backup()
            self.backup.backup_now("switch", force=False)

        # close current store
        try:
            self.store.release_session_locks(str(self.session.get("session_id", "")))
        except Exception:
            pass
        try:
            self.store.close()
        except Exception:
            pass

        # switch
        self.ws_mgr.set_current(ws_id)
        self.ws_id = ws_id
        self.ws = self.ws_mgr.get(ws_id) or self.ws_mgr.ensure_default()

        self.cfg_mgr = ConfigManager(self.ws_mgr.config_path(ws_id))
        self.cfg = self.cfg_mgr.load()

        self.store = Store(self.ws_mgr.db_path(ws_id))
        self.backup = BackupManager(self.ws_mgr, ws_id, self.store, retention_total=self.cfg.backup.retention_total)

        # Aggiorna riferimenti store/cfg nei tab modulari dopo switch workspace
        self._rebind_workspace_context_to_tabs(refresh_sw_tab=True)

        self._refresh_all()
        self._log_activity("WORKSPACE_SWITCH", status="OK", message=f"{old_ws} -> {self.ws_id}")
        info(f"Workspace attiva: {self.ws.name}")

    # ---------------- refresh
    def _refresh_all(self):
        self._set_ws_label()
        self._set_shared_root_label()
        
        # Refresh tab modulari
        if hasattr(self, 'tab_generatore_obj') and self.tab_generatore_obj:
            self.tab_generatore_obj.refresh_machines()
            self.tab_generatore_obj.refresh_groups()
        if hasattr(self, 'tab_codifica_obj') and self.tab_codifica_obj:
            self.tab_codifica_obj.refresh_machine_menus()
            self.tab_codifica_obj.refresh_vvv_menu()
        if hasattr(self, 'tab_gerarchia_obj') and self.tab_gerarchia_obj:
            self.tab_gerarchia_obj.refresh_tree()
        if hasattr(self, 'tab_monitor_obj') and self.tab_monitor_obj:
            self.tab_monitor_obj.refresh()
        if hasattr(self, 'tab_operativo_obj') and self.tab_operativo_obj:
            self.tab_operativo_obj.refresh_table()
            self.tab_operativo_obj.refresh_workflow()
        
        # Metodi ancora in app.py (da estrarre in futuri tab)
        # self._refresh_machines()  # -> gestito da tab_generatore_obj
        # self._refresh_groups()    # -> gestito da tab_generatore_obj
        # self._refresh_machine_menus()  # -> gestito da tab_codifica_obj
        # self._refresh_vvv_menu()  # -> gestito da tab_codifica_obj
        # self._refresh_hierarchy_tree()  # -> gestito da tab_gerarchia_obj
        # self._refresh_monitor_panel()  # -> gestito da tab_monitor_obj
        # self._refresh_rc_table()  # -> gestito da tab_operativo_obj
        # self._refresh_workflow_panel()  # -> gestito da tab_operativo_obj

    def _on_close(self):
        # Stop monitor auto-refresh
        if hasattr(self, 'tab_monitor_obj') and self.tab_monitor_obj:
            try:
                self.tab_monitor_obj.stop_auto_refresh()
            except Exception:
                pass
        
        # daily + switch-like backup on exit (solo se dirty)
        if self.cfg.backup.enabled:
            if self.cfg.backup.daily_enabled:
                self.backup.maybe_daily_backup()
            self.backup.backup_now("exit", force=False)
        try:
            self.store.release_session_locks(str(self.session.get("session_id", "")))
        except Exception:
            pass
        self._log_activity("APP_EXIT", status="OK", message="Desktop chiuso.")
        self.store.close()
        self.destroy()


if __name__ == "__main__":
    app = PDMApp()
    app.mainloop()
