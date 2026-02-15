from __future__ import annotations

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk
from typing import Callable, List, Sequence, Optional, Any


class SimpleTable(ttk.Frame):
    def __init__(self, master, columns: List[str], headings: List[str], on_double_click: Optional[Callable[[str], None]] = None, key_index: int = 0):
        super().__init__(master)
        self.columns = columns
        self.on_double_click = on_double_click
        self.key_index = key_index

        # Stile: aumenta testo righe (~ +50%)
        style = ttk.Style()
        try:
            base = tkfont.nametofont("TkDefaultFont")
            base_size = int(base.cget("size"))
            new_size = max(9, int(round(base_size * 1.5)))
            font_family = base.cget("family")
        except Exception:
            new_size = 14
            font_family = "Segoe UI"
        style.configure("PDM.Treeview", font=(font_family, new_size), rowheight=int(new_size * 2))
        style.configure("PDM.Treeview.Heading", font=(font_family, new_size, "bold"))

        self._sort_state = {}  # col -> bool (descending)

        self.tree = ttk.Treeview(self, columns=columns, show="headings", style="PDM.Treeview")
        # Colori riga per stato documento
        self.tree.tag_configure("state_wip", foreground="#000000")
        self.tree.tag_configure("state_in_rev", foreground="#B8860B")
        self.tree.tag_configure("state_rel", foreground="#0B5ED7")
        self.tree.tag_configure("state_obs", foreground="#7A7A7A")
        for c, h in zip(columns, headings):
            self.tree.heading(c, text=h, command=lambda col=c: self._on_sort(col))
            # colonne indicatori a sinistra: strette
            if c in ("m_ok", "d_ok"):
                self.tree.column(c, width=40, anchor=tk.CENTER, stretch=False)
            else:
                self.tree.column(c, width=140, anchor=tk.W)

        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        if on_double_click:
            self.tree.bind("<Double-1>", self._dbl)

    def _dbl(self, _evt):
        if not self.on_double_click:
            return
        sel = self.tree.selection()
        if not sel:
            return
        item = self.tree.item(sel[0])
        vals = item.get("values", [])
        if vals:
            k = vals[self.key_index] if len(vals) > self.key_index else vals[0]
            self.on_double_click(str(k))

    def _on_sort(self, col: str):
        desc = self._sort_state.get(col, False)
        self._sort_state[col] = not desc
        self.sort_by(col, descending=not desc)

    def sort_by(self, col: str, descending: bool = False):
        # Recupera valori
        items = list(self.tree.get_children(""))
        col_index = self.columns.index(col) if col in self.columns else 0

        def keyfunc(item_id: str):
            vals = self.tree.item(item_id).get("values", [])
            v = vals[col_index] if len(vals) > col_index else ""
            # tenta sort numerico
            try:
                s = str(v).strip()
                if s.isdigit():
                    return (0, int(s))
                # rev tipo '01'
                if len(s) <= 3 and s.replace('.', '').isdigit():
                    return (0, float(s))
            except Exception:
                pass
            return (1, str(v).lower())

        items.sort(key=keyfunc, reverse=descending)
        for idx, item_id in enumerate(items):
            self.tree.move(item_id, "", idx)

    

    def set_schema(self, columns: List[str], headings: List[str], key_index: Optional[int] = None):
        """Aggiorna colonne/intestazioni della tabella (utile per colonne dinamiche)."""
        if columns == self.columns and len(headings) == len(columns):
            if key_index is not None:
                self.key_index = int(key_index)
            return

        self.columns = list(columns)
        if key_index is not None:
            self.key_index = int(key_index)

        # reset sort state (colonne possono cambiare)
        self._sort_state = {}

        # aggiorna tree schema
        self.tree.configure(columns=self.columns)
        for c in self.columns:
            # se heading mancante, usa id
            h = c
            if headings and len(headings) == len(self.columns):
                h = headings[self.columns.index(c)]
            self.tree.heading(c, text=h, command=lambda col=c: self._on_sort(col))
            if c in ("m_ok", "d_ok"):
                self.tree.column(c, width=40, anchor=tk.CENTER, stretch=False)
            else:
                self.tree.column(c, width=140, anchor=tk.W, stretch=True)

    def clear(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

    def set_rows(self, rows: List[Any]):
        self.clear()
        for r in rows:
            values = r
            tags = ()
            if isinstance(r, dict):
                values = r.get("values", [])
                raw_tags = r.get("tags", ())
                if isinstance(raw_tags, str):
                    tags = (raw_tags,)
                else:
                    tags = tuple(raw_tags or ())
            self.tree.insert("", "end", values=list(values), tags=tags)


# Alias per compatibilità (in alcune release la tabella era chiamata Table)
class Table(SimpleTable):
    pass
