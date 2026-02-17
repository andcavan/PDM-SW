"""Tab Generatore - Gestione anagrafica macchine (MMM) e gruppi (GGGG)."""

from __future__ import annotations
from typing import TYPE_CHECKING
import tkinter as tk
import tkinter.font as tkfont
import customtkinter as ctk

from .base_tab import BaseTab, warn, ask

if TYPE_CHECKING:
    from pdm_sw.config import AppConfig
    from pdm_sw.store import Store


class TabGeneratore(BaseTab):
    """Tab per la gestione di macchine e gruppi macchina.
    
    Funzionalità:
    - Elenco macchine (MMM) con CRUD
    - Elenco gruppi (GGGG) per macchina selezionata con CRUD
    - Sincronizzazione selezione tra listbox e menu
    """
    
    def __init__(self, parent_frame, app, cfg: AppConfig, store: Store, session: dict):
        """Inizializza il tab Generatore.
        
        Args:
            parent_frame: Frame tkinter genitore dove costruire l'UI
            app: Riferimento all'applicazione principale
            cfg: Configurazione applicazione
            store: Store database
            session: Dati sessione utente
        """
        super().__init__(app, cfg, store, session)
        self.parent = parent_frame
        
        # Widget references
        self.machine_list = None
        self.group_list = None
        self.mmm_new = None
        self.mmm_name_new = None
        self.gggg_new = None
        self.gggg_name_new = None
        self.group_mmm_var = None
        self.group_mmm_menu = None
        
        self._build_ui()
    
    def _build_ui(self):
        """Costruisce l'interfaccia del tab."""
        outer = ctk.CTkFrame(self.parent)
        outer.pack(fill="both", expand=True, padx=10, pady=10)

        left = ctk.CTkFrame(outer)
        right = ctk.CTkFrame(outer)
        left.pack(side="left", fill="both", expand=True, padx=(0, 8))
        right.pack(side="left", fill="both", expand=True, padx=(8, 0))

        # === MACCHINE (MMM) ===
        ctk.CTkLabel(left, text="Macchine (MMM)", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=(10, 6))
        mlist_frame = ctk.CTkFrame(left)
        mlist_frame.pack(fill="both", expand=True, padx=10, pady=6)
        
        try:
            base = tkfont.nametofont("TkDefaultFont")
            base_size = int(base.cget("size"))
            lb_size = max(9, int(round(base_size * 1.5)))
            lb_family = base.cget("family")
            _lb_font = (lb_family, lb_size)
        except Exception:
            _lb_font = ("Segoe UI", 14)
        
        self.machine_list = tk.Listbox(mlist_frame, height=12, font=_lb_font, exportselection=False)
        m_scroll = tk.Scrollbar(mlist_frame, orient="vertical", command=self.machine_list.yview)
        self.machine_list.configure(yscrollcommand=m_scroll.set)
        self.machine_list.pack(side="left", fill="both", expand=True)
        m_scroll.pack(side="right", fill="y")
        self.machine_list.bind("<<ListboxSelect>>", lambda e: self._on_machine_list_selected())

        add_m = ctk.CTkFrame(left, fg_color="transparent")
        add_m.pack(fill="x", padx=10, pady=(6, 10))
        self.mmm_new = tk.StringVar()
        self.mmm_name_new = tk.StringVar()
        ctk.CTkEntry(add_m, placeholder_text="MMM", width=80, textvariable=self.mmm_new).pack(side="left", padx=(0, 6))
        ctk.CTkEntry(add_m, placeholder_text="Nome macchina", textvariable=self.mmm_name_new).pack(side="left", fill="x", expand=True, padx=6)
        ctk.CTkButton(add_m, text="Aggiungi", width=90, command=self._add_machine).pack(side="left", padx=6)
        ctk.CTkButton(add_m, text="Modifica", width=90, command=self._edit_machine_desc).pack(side="left", padx=6)
        ctk.CTkButton(add_m, text="Elimina", width=90, command=self._del_machine).pack(side="left", padx=6)

        # === GRUPPI (GGGG) ===
        # Selezione esplicita macchina per gruppi
        grp_hdr = ctk.CTkFrame(right, fg_color="transparent")
        grp_hdr.pack(fill="x", padx=10, pady=(10, 0))
        ctk.CTkLabel(grp_hdr, text="Macchina:").pack(side="left", padx=(0, 6))
        self.group_mmm_var = tk.StringVar(value="")
        self.group_mmm_menu = ctk.CTkOptionMenu(grp_hdr, variable=self.group_mmm_var, values=[""], width=140, command=lambda _=None: self._on_group_machine_selected())
        self.group_mmm_menu.pack(side="left")

        ctk.CTkLabel(right, text="Gruppi (GGGG) per macchina selezionata", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", padx=10, pady=(6, 6))
        glist_frame = ctk.CTkFrame(right)
        glist_frame.pack(fill="both", expand=True, padx=10, pady=6)
        self.group_list = tk.Listbox(glist_frame, height=12, font=_lb_font, exportselection=False)
        g_scroll = tk.Scrollbar(glist_frame, orient="vertical", command=self.group_list.yview)
        self.group_list.configure(yscrollcommand=g_scroll.set)
        self.group_list.pack(side="left", fill="both", expand=True)
        g_scroll.pack(side="right", fill="y")

        add_g = ctk.CTkFrame(right, fg_color="transparent")
        add_g.pack(fill="x", padx=10, pady=(6, 10))
        self.gggg_new = tk.StringVar()
        self.gggg_name_new = tk.StringVar()
        ctk.CTkEntry(add_g, placeholder_text="GGGG", width=90, textvariable=self.gggg_new).pack(side="left", padx=(0, 6))
        ctk.CTkEntry(add_g, placeholder_text="Nome gruppo", textvariable=self.gggg_name_new).pack(side="left", fill="x", expand=True, padx=6)
        ctk.CTkButton(add_g, text="Aggiungi", width=90, command=self._add_group).pack(side="left", padx=6)
        ctk.CTkButton(add_g, text="Modifica", width=90, command=self._edit_group_desc).pack(side="left", padx=6)
        ctk.CTkButton(add_g, text="Elimina", width=90, command=self._del_group).pack(side="left", padx=6)
    
    # === MACHINE HANDLERS ===
    
    def _on_machine_list_selected(self):
        """Gestisce selezione nella listbox macchine."""
        mmm = self._selected_mmm()
        if self.group_mmm_var:
            self.group_mmm_var.set(mmm)
        self.refresh_groups()
    
    def _selected_mmm(self) -> str:
        """Ottiene il codice macchina selezionato.
        
        Returns:
            Codice MMM selezionato (stringa vuota se nessuna selezione)
        """
        sel = self.machine_list.curselection()
        if sel:
            val = self.machine_list.get(sel[0])
            return val.split(" ")[0].strip()
        # fallback: usa il menu macchina quando la listbox non ha selezione attiva
        if self.group_mmm_var:
            return (self.group_mmm_var.get() or "").strip().upper()
        return ""
    
    def _add_machine(self):
        """Aggiunge una nuova macchina."""
        mmm = self._validate_segment_strict("MMM", self.mmm_new.get(), "MMM")
        if not mmm:
            warn("Inserisci MMM.")
            return
        name = self._require_desc_upper(self.mmm_name_new.get(), what="descrizione macchina")
        if name is None:
            return
        self.store.add_machine(mmm, name)
        self.mmm_new.set("")
        self.mmm_name_new.set("")
        self.refresh_machines()
        self.app._refresh_machine_menus()
        self.app._refresh_hierarchy_tree()
    
    def _edit_machine_desc(self):
        """Modifica la descrizione di una macchina esistente."""
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
        self.refresh_machines()
        if self.group_mmm_var:
            self.group_mmm_var.set(mmm)
        self._on_group_machine_selected()
        self.app._refresh_machine_menus()
        self.app._refresh_hierarchy_tree()
    
    def _del_machine(self):
        """Elimina una macchina."""
        mmm = self._selected_mmm()
        if not mmm:
            return
        if not ask(f"Eliminare macchina {mmm}?"):
            return
        self.store.delete_machine(mmm)
        self.refresh_machines()
        self.group_list.delete(0, tk.END)
        self.app._refresh_machine_menus()
        self.app._refresh_hierarchy_tree()
    
    def refresh_machines(self):
        """Aggiorna la lista delle macchine."""
        self.machine_list.delete(0, tk.END)
        machines = []
        for mmm, name in self.store.list_machines():
            machines.append(mmm)
            self.machine_list.insert(tk.END, f"{mmm} - {name}")
        if self.group_mmm_menu:
            self.group_mmm_menu.configure(values=(machines if machines else [""]))
            if self.group_mmm_var.get() not in machines:
                self.group_mmm_var.set(machines[0] if machines else "")
    
    # === GROUP HANDLERS ===
    
    def _selected_gggg(self) -> str:
        """Ottiene il codice gruppo selezionato.
        
        Returns:
            Codice GGGG selezionato (stringa vuota se nessuna selezione)
        """
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
    
    def _on_group_machine_selected(self):
        """Gestisce cambio macchina nel menu a tendina della sezione gruppi."""
        mmm = (self.group_mmm_var.get() or "").strip().upper() if self.group_mmm_var else ""
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
        self.refresh_groups()
        self.app._refresh_group_menu()
    
    def _add_group(self):
        """Aggiunge un nuovo gruppo."""
        mmm = (self.group_mmm_var.get().strip() if self.group_mmm_var else self._selected_mmm())
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
        self.gggg_new.set("")
        self.gggg_name_new.set("")
        self.refresh_groups()
        self.app._refresh_group_menu()
        self.app._refresh_hierarchy_tree()
    
    def _edit_group_desc(self):
        """Modifica la descrizione di un gruppo esistente."""
        mmm = (self.group_mmm_var.get().strip().upper() if self.group_mmm_var else self._selected_mmm())
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
        self.refresh_groups()
        self.app._refresh_group_menu()
        self.app._refresh_hierarchy_tree()
    
    def _del_group(self):
        """Elimina un gruppo."""
        mmm = (self.group_mmm_var.get().strip().upper() if self.group_mmm_var else self._selected_mmm())
        sel = self.group_list.curselection()
        if not mmm or not sel:
            return
        gggg = self.group_list.get(sel[0]).split(" ")[0].strip()
        if not ask(f"Eliminare gruppo {mmm}/{gggg}?"):
            return
        self.store.delete_group(mmm, gggg)
        self.refresh_groups()
        self.app._refresh_group_menu()
        self.app._refresh_hierarchy_tree()
    
    def refresh_groups(self):
        """Aggiorna la lista dei gruppi per la macchina selezionata."""
        mmm = self._selected_mmm()
        self.group_list.delete(0, tk.END)
        if not mmm:
            return
        for gggg, name in self.store.list_groups(mmm):
            self.group_list.insert(tk.END, f"{gggg} - {name}")
