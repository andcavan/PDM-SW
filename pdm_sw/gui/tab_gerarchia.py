# -*- coding: utf-8 -*-
"""
Tab Gerarchia
Visualizza struttura gerarchica MMM -> GGGG -> codici
"""
from __future__ import annotations
from typing import TYPE_CHECKING
import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk
from pathlib import Path
from collections import defaultdict
import customtkinter as ctk
from .base_tab import BaseTab

if TYPE_CHECKING:
    from pdm_sw.models import Document


class TabGerarchia(BaseTab):
    """Tab con treeview gerarchica MMM -> GGGG -> codici."""

    def __init__(self, parent_frame, app, cfg, store, session):
        super().__init__(app, cfg, store, session)
        self.root = parent_frame
        self.hierarchy_tree = None
        self.hierarchy_include_obs_var = None
        self._build_ui()

    def _build_ui(self):
        """Costruisce UI con treeview gerarchica."""
        actions = ctk.CTkFrame(self.root, fg_color="#F2D65C")
        actions.pack(fill="x", padx=10, pady=(10, 6))

        ctk.CTkButton(
            actions,
            text="AGGIORNA",
            width=120,
            command=self.refresh_tree
        ).pack(side="left", padx=6, pady=6)
        
        ctk.CTkButton(
            actions,
            text="REPORT GENERALE",
            width=160,
            command=self._generate_hierarchy_report
        ).pack(side="left", padx=6, pady=6)
        
        self.hierarchy_include_obs_var = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            actions,
            text="Mostra OBS",
            variable=self.hierarchy_include_obs_var,
            command=self.refresh_tree,
        ).pack(side="left", padx=(12, 6), pady=6)

        ctk.CTkLabel(
            actions,
            text="Struttura: MMM -> GGGG -> codici PART/ASSY",
        ).pack(side="right", padx=10, pady=6)

        # Tree frame
        tree_frame = ctk.CTkFrame(self.root)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Stile treeview
        style = ttk.Style()
        try:
            base = tkfont.nametofont("TkDefaultFont")
            base_size = int(base.cget("size"))
            tree_size = max(10, int(round(base_size * 1.45)))
            tree_family = base.cget("family")
        except Exception:
            tree_size = 15
            tree_family = "Segoe UI"
        
        style.configure(
            "PDM.Hierarchy.Treeview",
            font=(tree_family, tree_size),
            rowheight=int(tree_size * 2),
        )

        # Treeview
        self.hierarchy_tree = ttk.Treeview(
            tree_frame,
            show="tree",
            style="PDM.Hierarchy.Treeview"
        )
        
        self.hierarchy_tree.tag_configure("part_node", foreground="#0B5ED7")
        self.hierarchy_tree.tag_configure("assy_node", foreground="#2E7D32")
        self.hierarchy_tree.tag_configure("machine_node", foreground="#E67E22", font=(tree_family, tree_size, "bold"))
        self.hierarchy_tree.tag_configure("group_node", foreground="#3498DB", font=(tree_family, tree_size, "bold"))
        
        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.hierarchy_tree.yview)
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.hierarchy_tree.xview)
        self.hierarchy_tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.hierarchy_tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.hierarchy_tree.bind("<Double-1>", self._on_double_click)

    def refresh_tree(self):
        """Aggiorna il contenuto del treeview."""
        if not self.hierarchy_tree:
            return

        tree = self.hierarchy_tree
        current_nodes = tree.get_children()
        if current_nodes:
            tree.delete(*current_nodes)

        include_obs = bool(self.hierarchy_include_obs_var.get()) if self.hierarchy_include_obs_var else False

        machines_raw = self.store.list_machines()
        machine_names: dict[str, str] = {mmm: name for mmm, name in machines_raw}

        groups_by_machine: dict[str, dict[str, str]] = {}
        for mmm, _ in machines_raw:
            groups_by_machine[mmm] = {gggg: g_name for gggg, g_name in self.store.list_groups(mmm)}

        docs = self.store.list_documents(include_obs=include_obs)
        docs_by_pair: dict[tuple[str, str], list] = defaultdict(list)
        docs_machine: dict[str, list] = defaultdict(list)  # MACHINE per MMM

        for d in docs:
            if d.doc_type == "MACHINE":
                docs_machine[d.mmm].append(d)
            else:
                docs_by_pair[(d.mmm, d.gggg)].append(d)
            machine_names.setdefault(d.mmm, "")
            groups_by_machine.setdefault(d.mmm, {})
            if d.gggg:
                groups_by_machine[d.mmm].setdefault(d.gggg, "")

        if not machine_names:
            tree.insert("", "end", text="(nessun MMM disponibile)")
            return

        for mmm in sorted(machine_names.keys()):
            m_name = (machine_names.get(mmm) or "").strip()
            m_text = f"{mmm} - {m_name}" if m_name else mmm
            m_node = tree.insert("", "end", text=m_text, open=True)

            # Mostra MACHINE versions sotto MMM
            machine_docs = docs_machine.get(mmm, [])
            if machine_docs:
                machine_docs.sort(key=lambda d: d.code)
                for d in machine_docs:
                    model_path = d.best_model_path_for_state()
                    drw_path = d.best_drw_path_for_state()
                    d_node = tree.insert(
                        m_node,
                        "end",
                        text=f"📦 {d.code}",
                        values=(d.code,),
                        tags=("machine_node",)
                    )
                    tree.insert(d_node, "end", text=f"DESC: {d.description}")
                    tree.insert(
                        d_node,
                        "end",
                        text=f"MODEL ({d.state}): {model_path if model_path else 'NON ASSOCIATO'}"
                    )
                    tree.insert(
                        d_node,
                        "end",
                        text=f"DRW ({d.state}): {drw_path if drw_path else 'NON ASSOCIATO'}"
                    )

            group_map = groups_by_machine.get(mmm, {})
            gggg_keys = sorted(group_map.keys())
            if not gggg_keys:
                if not machine_docs:
                    tree.insert(m_node, "end", text="(nessun GGGG o versione macchina)")
                continue

            for gggg in gggg_keys:
                g_name = (group_map.get(gggg) or "").strip()
                dlist = docs_by_pair.get((mmm, gggg), [])
                dlist.sort(key=lambda d: (0 if d.doc_type == "PART" else (1 if d.doc_type == "ASSY" else 2), d.code))

                part_count = sum(1 for d in dlist if d.doc_type == "PART")
                assy_count = sum(1 for d in dlist if d.doc_type == "ASSY")
                group_count = sum(1 for d in dlist if d.doc_type == "GROUP")

                g_base = f"{gggg} - {g_name}" if g_name else gggg
                g_text = f"{g_base} (GROUP:{group_count} PART:{part_count} ASSY:{assy_count})"
                g_node = tree.insert(m_node, "end", text=g_text, open=True)

                if not dlist:
                    tree.insert(g_node, "end", text="(nessun codice)")
                    continue

                for d in dlist:
                    model_path = d.best_model_path_for_state()
                    drw_path = d.best_drw_path_for_state()
                    
                    if d.doc_type == "PART":
                        tags = ("part_node",)
                        icon = ""
                    elif d.doc_type == "ASSY":
                        tags = ("assy_node",)
                        icon = ""
                    elif d.doc_type == "GROUP":
                        tags = ("group_node",)
                        icon = "📁 "
                    else:
                        tags = ()
                        icon = ""
                    
                    d_node = tree.insert(
                        g_node,
                        "end",
                        text=f"{icon}{d.code}",
                        values=(d.code,),
                        tags=tags
                    )
                    
                    if d.doc_type == "GROUP":
                        tree.insert(d_node, "end", text=f"DESC: {d.description}")
                    
                    tree.insert(
                        d_node,
                        "end",
                        text=f"MODEL ({d.state}): {model_path if model_path else 'NON ASSOCIATO'}"
                    )
                    tree.insert(
                        d_node,
                        "end",
                        text=f"DRW ({d.state}): {drw_path if drw_path else 'NON ASSOCIATO'}"
                    )

    def _on_double_click(self, _evt=None):
        """Gestisce doppio click su nodo tree: apre codice in tab Operativo."""
        if not self.hierarchy_tree:
            return
        try:
            sel = self.hierarchy_tree.selection()
            if not sel:
                return
            values = self.hierarchy_tree.item(sel[0]).get("values", [])
            code = str(values[0]).strip() if values else ""
            if not code:
                return
            
            # Imposta il codice in workflow e passa a tab Operativo
            if hasattr(self.app, 'wf_code_var'):
                self.app.wf_code_var.set(code)
            
            try:
                self.app.tabs.set("Operativo")
            except Exception:
                pass
            
            if hasattr(self.app, '_refresh_workflow_panel'):
                self.app._refresh_workflow_panel()
        except Exception:
            return

    def _generate_hierarchy_report(self):
        """Delega al mixin ReportMixin per generare report."""
        if hasattr(self.app, '_generate_hierarchy_report'):
            self.app._generate_hierarchy_report()
