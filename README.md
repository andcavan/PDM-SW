# PDM SolidWorks (Python + CustomTkinter) — Workspace Edition

## Avvio rapido (Windows)
```bat
py -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe app.py
```

## Workspaces
- Ogni **WORKSPACE** è indipendente e vive dentro `WORKSPACES/`.
- Barra in alto: **Nome + Descrizione** sempre visibili e pulsanti:
  - CAMBIA / CREA / CANCELLA / COPIA WORKSPACE
- Ogni workspace contiene `config.json`, `pdm.db`, `backups/`.

## Backup DB (per workspace)
Default:
- backup automatico **ON**
- backup prima di cambiare workspace (se DB modificato)
- backup su eventi critici workflow (Release / Approve / Obsolete)
- retention: ultime 30 copie

## SolidWorks (COM)
- Per creare file da template serve SolidWorks installato + `pywin32`.
- Il DRW (se creato) ha sempre lo stesso nome del modello (estensione diversa).
