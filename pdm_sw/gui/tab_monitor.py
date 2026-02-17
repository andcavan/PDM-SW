"""
Tab Monitor - Visualizzazione lock attivi e activity log con auto-refresh.

Espone:
- ActivityLog viewer con refresh automatico ogni 5 secondi
- Lock attivi con tempo residuo prima della scadenza
- Filtro limite eventi (100/200/500/1000)
- Checkbox per abilitare/disabilitare auto-refresh
"""
import tkinter as tk
from datetime import datetime

import customtkinter as ctk

from .base_tab import BaseTab
from pdm_sw.ui.table import Table


class TabMonitor(BaseTab):
    """
    Tab Monitor: visualizzazione lock attivi e activity log.
    
    - Lock attivi: mostra i lock in corso con tempo residuo
    - Activity log: eventi recenti filtrati per limite
    - Auto-refresh: polling ogni 5 secondi se abilitato
    """
    
    def __init__(self, parent_frame, app, cfg, store, session):
        """
        Inizializza il tab Monitor.
        
        Args:
            parent_frame: frame genitore in cui costruire la UI
            app: riferimento all'app principale
            cfg: configurazione workspace
            store: store SQLite
            session: sessione utente
        """
        super().__init__(app, cfg, store, session)
        self.root = parent_frame
        self.monitor_auto_var = None
        self.monitor_limit_var = None
        self.monitor_summary_var = None
        self.monitor_lock_table = None
        self.monitor_activity_table = None
        self.monitor_after_id = None
        self._build_ui()
    
    def _build_ui(self):
        """Costruisce la UI del tab Monitor."""
        tab = self.root
        
        self.monitor_auto_var = tk.BooleanVar(value=True)
        self.monitor_limit_var = tk.StringVar(value="200")
        self.monitor_after_id = None
        
        actions = ctk.CTkFrame(tab, fg_color="#EAF2FF")
        actions.pack(fill="x", padx=10, pady=(10, 6))
        
        ctk.CTkButton(
            actions,
            text="AGGIORNA ORA",
            width=140,
            command=self.refresh
        ).pack(side="left", padx=6, pady=6)
        
        ctk.CTkCheckBox(
            actions,
            text="Auto refresh 5s",
            variable=self.monitor_auto_var,
            command=self.refresh
        ).pack(side="left", padx=12, pady=6)
        
        ctk.CTkLabel(
            actions,
            text="Attivita recenti:"
        ).pack(side="left", padx=(12, 4), pady=6)
        
        ctk.CTkComboBox(
            actions,
            variable=self.monitor_limit_var,
            values=["100", "200", "500", "1000"],
            width=90,
            command=lambda _=None: self.refresh()
        ).pack(side="left", padx=4, pady=6)
        
        self.monitor_summary_var = tk.StringVar(value="Lock attivi: 0 | Eventi: 0")
        ctk.CTkLabel(
            actions,
            textvariable=self.monitor_summary_var,
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(side="right", padx=10, pady=6)
        
        # Lock attivi
        locks_box = ctk.CTkFrame(tab)
        locks_box.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        
        ctk.CTkLabel(
            locks_box,
            text="Lock Attivi",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=8, pady=(8, 4))
        
        lock_cols = ["code", "owner_user", "owner_host", "acquired_at", "updated_at", "expires_at", "remaining_min"]
        lock_heads = ["CODICE", "UTENTE", "HOST", "ACQUISITO", "AGGIORNATO", "SCADENZA", "MIN RESIDUI"]
        self.monitor_lock_table = Table(locks_box, columns=lock_cols, headings=lock_heads, key_index=0)
        self.monitor_lock_table.pack(fill="both", expand=True, padx=6, pady=(0, 8))
        self.monitor_lock_table.tree.tag_configure("lock_mine", foreground="#0B5ED7")
        
        # Activity log
        activity_box = ctk.CTkFrame(tab)
        activity_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        ctk.CTkLabel(
            activity_box,
            text="Activity Log",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=8, pady=(8, 4))
        
        act_cols = ["created_at", "action", "status", "code", "user_display", "host", "file_path", "message"]
        act_heads = ["DATA/ORA", "AZIONE", "ESITO", "CODICE", "UTENTE", "HOST", "FILE", "MESSAGGIO"]
        self.monitor_activity_table = Table(activity_box, columns=act_cols, headings=act_heads, key_index=3)
        self.monitor_activity_table.pack(fill="both", expand=True, padx=6, pady=(0, 8))
        
        try:
            self.monitor_activity_table.tree.column("file_path", width=360, stretch=True)
            self.monitor_activity_table.tree.column("message", width=260, stretch=True)
        except Exception:
            pass
        
        self.monitor_activity_table.tree.tag_configure("act_error", foreground="#B91C1C")
        self.monitor_activity_table.tree.tag_configure("act_warn", foreground="#B8860B")
        self.monitor_activity_table.tree.tag_configure("act_ok", foreground="#166534")
    
    def refresh(self):
        """
        Aggiorna i dati di lock attivi e activity log.
        
        Chiamato:
        - Dal bottone AGGIORNA ORA
        - Dall'auto-refresh ogni 5 secondi (se abilitato)
        - Dal cambio di workspace
        - Dall'app._refresh_all()
        """
        if not self.monitor_lock_table or not self.monitor_activity_table:
            return
        
        try:
            limit = int((self.monitor_limit_var.get() or "200").strip())
        except Exception:
            limit = 200
        limit = min(1000, max(50, limit))
        
        my_session = str(self.session.get("session_id", ""))
        now = datetime.now()
        
        # Lock attivi
        try:
            locks = self.store.list_active_locks(limit=500)
        except Exception:
            locks = []
        
        lock_rows = []
        for lk in locks:
            exp = str(lk.get("expires_at", ""))
            rem = ""
            try:
                rem_min = max(0.0, (datetime.fromisoformat(exp) - now).total_seconds() / 60.0)
                rem = f"{rem_min:.1f}"
            except Exception:
                rem = ""
            
            vals = [
                str(lk.get("code", "")),
                str(lk.get("owner_user", "")),
                str(lk.get("owner_host", "")),
                str(lk.get("acquired_at", "")).replace("T", " "),
                str(lk.get("updated_at", "")).replace("T", " "),
                exp.replace("T", " "),
                rem,
            ]
            tags = ("lock_mine",) if str(lk.get("owner_session", "")) == my_session else ()
            lock_rows.append({"values": vals, "tags": tags})
        
        self.monitor_lock_table.set_rows(lock_rows)
        
        # Activity log
        try:
            activities = self.store.list_recent_activity(limit=limit)
        except Exception:
            activities = []
        
        act_rows = []
        for a in activities:
            status = str(a.get("status", "")).strip().upper()
            details = a.get("details", {}) if isinstance(a.get("details", {}), dict) else {}
            file_path = str(details.get("path", "") or "")
            
            tag = ""
            if status in ("ERROR", "KO"):
                tag = "act_error"
            elif status in ("WARN", "WARNING", "LOCKED"):
                tag = "act_warn"
            elif status == "OK":
                tag = "act_ok"
            
            vals = [
                str(a.get("created_at", "")).replace("T", " "),
                str(a.get("action", "")),
                status,
                str(a.get("code", "")),
                str(a.get("user_display", "")),
                str(a.get("host", "")),
                file_path,
                str(a.get("message", "")),
            ]
            act_rows.append({"values": vals, "tags": ((tag,) if tag else ())})
        
        self.monitor_activity_table.set_rows(act_rows)
        
        self.monitor_summary_var.set(f"Lock attivi: {len(lock_rows)} | Eventi: {len(act_rows)}")
        
        # Auto-refresh scheduling
        if getattr(self, "monitor_after_id", None):
            try:
                self.app.after_cancel(self.monitor_after_id)
            except Exception:
                pass
            self.monitor_after_id = None
        
        if bool(self.monitor_auto_var.get()):
            self.monitor_after_id = self.app.after(5000, self.refresh)
    
    def stop_auto_refresh(self):
        """Ferma l'auto-refresh (chiamato alla chiusura del tab/app)."""
        if getattr(self, "monitor_after_id", None):
            try:
                self.app.after_cancel(self.monitor_after_id)
            except Exception:
                pass
            self.monitor_after_id = None
