# -*- coding: utf-8 -*-
"""
Tab Manuale
Contiene la guida rapida dell'applicazione
"""
import customtkinter as ctk
from .base_tab import BaseTab


class TabManuale(BaseTab):
    """Tab con manuale d'uso rapido."""

    def __init__(self, parent_frame, app, cfg, store, session):
        super().__init__(app, cfg, store, session)
        self.root = parent_frame
        self._build_ui()

    def _build_ui(self):
        # Sfondo giallo chiaro
        try:
            self.root.configure(fg_color="#FFF4B8")
        except Exception:
            pass

        ctk.CTkLabel(
            self.root,
            text="Manuale Rapido (Rev v50.2)",
            font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(anchor="w", padx=12, pady=(12, 8))

        ctk.CTkLabel(
            self.root,
            text="I dettagli completi sono nel file README.md.",
            text_color="#555555",
        ).pack(anchor="w", padx=12, pady=(0, 8))

        box = ctk.CTkTextbox(self.root)
        box.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        manual_text = (
            "PANORAMICA FUNZIONALITA\n"
            "=======================\n"
            "1) Workspace\n"
            "- Usa il pulsante WORKSPACE... in alto per aprire gli strumenti workspace.\n"
            "- Da li puoi cambiare/creare/copiare/cancellare workspace e scegliere la cartella condivisa.\n"
            "- Ogni workspace ha configurazione, database e backup separati.\n"
            "- Puoi scegliere la CARTELLA CONDIVISA (default: cartella attuale).\n"
            "- La root condivisa attiva e mostrata in alto come SHARED: <percorso>.\n"
            "- Le funzioni di configurazione sono raccolte nella tab principale SETUP.\n"
            "\n"
            "2) Gestione codifica\n"
            "- Configura formato codice: MMM, GGGG, VVV e progressivo.\n"
            "- Definisci separatori, lunghezze e regole di validazione.\n"
            "\n"
            "3) Generatore codici\n"
            "- Gestisci elenco macchine (MMM) e gruppi (GGGG).\n"
            "\n"
            "4) SolidWorks\n"
            "- Imposta archivio e template PART/ASSY/DRW.\n"
            "- Configura mappatura proprieta PDM -> SW e lettura SW -> PDM.\n"
            "- Pubblica la macro bootstrap per la workspace corrente.\n"
            "\n"
            "5) Codifica\n"
            "- Crea solo codice (WIP) oppure codice + file da template.\n"
            "- Importa file esistenti (.sldprt/.sldasm/.slddrw) nel WIP.\n"
            "\n"
            "6) Gerarchia\n"
            "- Vista ad albero MMM -> GGGG -> CODICE.\n"
            "\n"
            "7) Operativo\n"
            "- Tab unica per Ricerca&Consultazione + Workflow.\n"
            "- Sinistra: filtri, tabella risultati, apertura MODEL/DRW, creazione file mancanti e comandi FORZA PDM->SW / SW->PDM.\n"
            "- Destra: pannello workflow con transizioni stato e dettagli del codice selezionato.\n"
            "- Il pulsante INVIA A WORKFLOW e nel pannello workflow (in alto).\n"
            "- Puoi regolare la larghezza del pannello workflow trascinando il separatore centrale (salvata in locale).\n"
            "- Transizioni stato: WIP -> REL -> IN_REV -> REL e OBS.\n"
            "- Ripristino da OBS allo stato precedente se disponibile.\n"
            "- Ogni cambio stato richiede una nota obbligatoria (minimo 3 caratteri).\n"
            "- Le note sono salvate con data/ora, tipo evento e stati (da -> a).\n"
            "- Ogni transizione registra operazioni file in WORKSPACES\\<workspace>\\LOGS\\workflow.log.\n"
            "\n"
            "8) Monitor\n"
            "- Vista LOCK ATTIVI: codice, utente, host, scadenza lock e minuti residui.\n"
            "- Vista ACTIVITY LOG: data/ora, azione, esito, codice, utente, host, file e messaggio.\n"
            "- Auto refresh ogni 5 secondi (disattivabile) e filtro numero eventi recenti.\n"
            "\n"
            "9) Multiutente (3-6 utenti)\n"
            "- I lock sono condivisi nel DB workspace e proteggono le transizioni workflow.\n"
            "- Se un utente ha lock su un codice, gli altri vedono il blocco e non possono eseguire la stessa transizione.\n"
            "- I lock vengono rilasciati a fine operazione o alla chiusura sessione; in caso crash scadono a timeout.\n"
            "- Tracciamento aperture file (livello 1): log da tab Ricerca&Consultazione (doppio click/APRI MODELLO/APRI DRW).\n"
            "\n"
            "INSTALLAZIONE MACRO SOLIDWORKS (COMPLETA)\n"
            "=========================================\n"
            "Prerequisiti:\n"
            "- SolidWorks installato.\n"
            "- Tab SolidWorks compilata (Archivio + template) e salvata.\n"
            "\n"
            "Passi:\n"
            "1) Vai in tab SolidWorks e premi: PUBBLICA MACRO SOLIDWORKS.\n"
            "2) Apri il file istruzioni generato in:\n"
            "   SW_MACROS\\INSTALL_MACRO_<workspace_id>.txt\n"
            "3) In SolidWorks: Strumenti > Macro > Nuova...\n"
            "4) Salva la macro in:\n"
            "   SW_MACROS\\PDM_SW_BOOTSTRAP_<workspace_id>.swp\n"
            "5) Nell'editor VBA: File > Import File... e importa:\n"
            "   SW_MACROS\\PDM_SW_BOOTSTRAP_<workspace_id>.bas\n"
            "6) Salva e chiudi VBA.\n"
            "7) (Consigliato) Aggiungi pulsante toolbar:\n"
            "   Strumenti > Personalizza > Comandi > Macro > Esegui Macro.\n"
            "\n"
            "Build EXE payload (opzionale ma consigliato):\n"
            "1) Apri:\n"
            "   WORKSPACES\\<workspace_folder>\\macros\\payload\n"
            "2) Esegui:\n"
            "   build_payload_exe.bat\n"
            "3) Verifica presenza di PDM_SW_PAYLOAD.exe nella cartella payload.\n"
            "\n"
            "Diagnostica macro:\n"
            "- Log bootstrap:\n"
            "  SW_CACHE\\<workspace_id>\\payload\\bootstrap.log\n"
            "- Log payload:\n"
            "  SW_CACHE\\<workspace_id>\\payload\\payload.log\n"
        )

        box.insert("1.0", manual_text)
        box.configure(state="disabled")
