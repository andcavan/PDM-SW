"""
Helper functions comuni per la GUI.
"""
from __future__ import annotations

from tkinter import messagebox


def warn(msg: str) -> None:
    """Mostra warning dialog."""
    messagebox.showwarning("PDM", msg)


def info(msg: str) -> None:
    """Mostra info dialog."""
    messagebox.showinfo("PDM", msg)


def ask(msg: str) -> bool:
    """Mostra yes/no dialog."""
    return messagebox.askyesno("PDM", msg)


def error(msg: str) -> None:
    """Mostra error dialog."""
    messagebox.showerror("PDM", msg)
