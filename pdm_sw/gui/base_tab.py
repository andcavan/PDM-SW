"""Classe base per tab modulari dell'applicazione PDM-SW."""

from __future__ import annotations
from typing import TYPE_CHECKING
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk

if TYPE_CHECKING:
    from pdm_sw.config import AppConfig
    from pdm_sw.store import Store


def warn(msg: str) -> None:
    """Mostra un avviso all'utente."""
    messagebox.showwarning("PDM", msg)


def info(msg: str) -> None:
    """Mostra un'informazione all'utente."""
    messagebox.showinfo("PDM", msg)


def ask(msg: str) -> bool:
    """Richiede conferma all'utente."""
    return messagebox.askyesno("PDM", msg)


class BaseTab:
    """Classe base per tutti i tab dell'applicazione.
    
    Fornisce accesso a:
    - cfg: Configurazione applicazione
    - store: Database store
    - session: Dati sessione utente
    - app: Riferimento alla finestra principale
    
    Include metodi helper comuni per validazione e dialoghi.
    """
    
    def __init__(self, app, cfg: AppConfig, store: Store, session: dict):
        """Inizializza il tab base.
        
        Args:
            app: Riferimento alla finestra principale (PDMApp)
            cfg: Configurazione applicazione
            store: Store database
            session: Dizionario con dati sessione utente
        """
        self.app = app
        self.cfg = cfg
        self.store = store
        self.session = session
    
    def _validate_segment_strict(self, seg: str, value: str, what: str) -> str | None:
        """Valida un segmento secondo regole di Gestione Codifica.

        Regole:
        - UPPER/LOWER viene forzato sempre
        - Se lunghezza è impostata (in eccesso o difetto) -> errore
        - Se charset (ALPHA/NUM/ALNUM) non rispettato -> errore
        - Non applica padding/troncamenti né rimuove caratteri: o è valido o fallisce.
        
        Args:
            seg: Nome del segmento (es. "MMM", "GGGG")
            value: Valore da validare
            what: Descrizione per messaggi di errore
            
        Returns:
            Valore normalizzato se valido, None altrimenti
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
        """Valida e normalizza una descrizione.
        
        Args:
            desc: Testo descrizione
            what: Tipo di descrizione per messaggi di errore
            
        Returns:
            Descrizione normalizzata (uppercase) se valida, None altrimenti
        """
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
        """Mostra finestra di dialogo per inserimento testo lungo.
        
        Args:
            title: Titolo finestra
            prompt: Testo del prompt
            initial: Valore iniziale
            
        Returns:
            Testo inserito dall'utente, None se annullato
        """
        result: dict[str, str | None] = {"value": None}

        top = ctk.CTkToplevel(self.app)
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

        self.app.wait_window(top)
        return result["value"]
