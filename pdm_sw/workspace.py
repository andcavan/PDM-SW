from __future__ import annotations

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional
import json
import uuid
import re
from datetime import datetime


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def sanitize_name(name: str) -> str:
    """Rende un nome adatto a diventare parte del nome cartella 'id_nome'."""
    n = (name or "").strip()
    n = re.sub(r"\s+", "_", n)
    # consenti solo alfanumerici, underscore e dash
    n = re.sub(r"[^A-Za-z0-9_-]+", "", n)
    return n or "workspace"


@dataclass
class Workspace:
    id: str
    name: str
    description: str = ""
    # path è il nome cartella relativo (es. '8564ba90_ClienteX') oppure un vecchio path assoluto (legacy).
    path: str = ""
    created_at: str = ""
    updated_at: str = ""


class WorkspaceManager:
    """Workspace index lives inside WORKSPACES/; each workspace is fully independent."""

    def __init__(self, base_dir: Path):
        self.base_dir = Path(base_dir)
        self.base_dir.mkdir(parents=True, exist_ok=True)

        self.index_path = self.base_dir / "workspaces.json"
        self.current_path = self.base_dir / "current_workspace.txt"

        self._index: Dict[str, Workspace] = {}
        self._load_index()

    # ---------------------------
    # Internal helpers
    # ---------------------------

    def _resolve_ws_dir(self, ws: Workspace) -> Path:
        """Ritorna il path reale della workspace sotto base_dir.

        Supporta vecchi indici che salvavano path assoluti (portabilità quando si sposta C:\\PDM-SW).
        Se il path assoluto non esiste più, cade su base_dir/<nome_cartella>.
        """
        p = Path(ws.path or "")
        if not str(p).strip():
            folder = f"{ws.id}_{sanitize_name(ws.name)}"
            return self.base_dir / folder

        if p.is_absolute():
            if p.exists():
                return p
            return self.base_dir / p.name

        return self.base_dir / p


    def resolve_ws_dir(self, ws) -> Path:
        """Alias pubblico per compatibilità (macro_runtime v47.30+).
        Accetta Workspace oppure ws_id (str) e ritorna il path reale della cartella workspace.
        """
        if isinstance(ws, Workspace):
            return self._resolve_ws_dir(ws)
        ws_id = (ws or "").strip()
        w = self.get(ws_id)
        if not w:
            raise ValueError("Workspace non trovata")
        return self._resolve_ws_dir(w)

    def _ensure_ws_subdirs(self, ws_dir: Path) -> None:
        """Crea le sottocartelle standard della workspace (id_nome), se mancanti."""
        ws_dir = Path(ws_dir)
        (ws_dir / "macros" / "payload").mkdir(parents=True, exist_ok=True)
        (ws_dir / "macros" / "bootstrap").mkdir(parents=True, exist_ok=True)
        (ws_dir / "backups").mkdir(parents=True, exist_ok=True)

    def _save_index(self) -> None:
        payload = {"workspaces": {k: asdict(v) for k, v in self._index.items()}}
        self.index_path.write_text(
            json.dumps(payload, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )

    def _load_index(self) -> None:
        if self.index_path.exists():
            try:
                raw = json.loads(self.index_path.read_text(encoding="utf-8"))
            except Exception:
                raw = {}
            self._index = {
                k: Workspace(**{kk: vv for kk, vv in v.items() if kk in Workspace.__dataclass_fields__})
                for k, v in (raw.get("workspaces", {}) or {}).items()
            }
        else:
            self._index = {}
            self._save_index()

        # Migrazione: salva path relativo (folder) invece di path assoluto.
        changed = False
        for k, ws in list(self._index.items()):
            try:
                p = Path(ws.path or "")
                if p.is_absolute():
                    ws.path = p.name  # folder id_nome
                    changed = True
                else:
                    ws.path = str(p).strip() if ws.path else ws.path

                ws_dir = self._resolve_ws_dir(ws)
                ws_dir.mkdir(parents=True, exist_ok=True)
                self._ensure_ws_subdirs(ws_dir)

                self._index[k] = ws
            except Exception:
                pass

        if changed:
            self._save_index()

    # ---------------------------
    # Public API
    # ---------------------------

    def list(self) -> List[Workspace]:
        return sorted(self._index.values(), key=lambda w: (w.name or "").lower())

    def get(self, workspace_id: str) -> Optional[Workspace]:
        return self._index.get(workspace_id)

    def ensure_default(self) -> Workspace:
        cur = self.get(self.get_current_id())
        if cur:
            return cur
        if self._index:
            first = self.list()[0]
            self.set_current(first.id)
            return first
        ws = self.create("DEFAULT", "Workspace predefinita")
        self.set_current(ws.id)
        return ws

    def get_current_id(self) -> str:
        if self.current_path.exists():
            return (self.current_path.read_text(encoding="utf-8") or "").strip()
        return ""

    def set_current(self, ws_id: str) -> None:
        self.current_path.write_text(ws_id, encoding="utf-8")

    def create(self, name: str, description: str = "") -> Workspace:
        ws_id = uuid.uuid4().hex[:8]
        folder = f"{ws_id}_{sanitize_name(name)}"
        ws_dir = self.base_dir / folder
        ws_dir.mkdir(parents=True, exist_ok=True)
        self._ensure_ws_subdirs(ws_dir)

        ws = Workspace(
            id=ws_id,
            name=name.strip(),
            description=(description or "").strip(),
            path=folder,
            created_at=_now_iso(),
            updated_at=_now_iso(),
        )
        self._index[ws_id] = ws
        self._save_index()
        return ws

    def copy(self, src_id: str, name: str, description: str = "", copy_db: bool = True) -> Workspace:
        src = self.get(src_id)
        if not src:
            raise ValueError("Workspace sorgente non trovata")

        ws = self.create(name, description)
        dst_dir = self.workspace_dir(ws.id)
        src_dir = self.workspace_dir(src_id)

        # copia config
        try:
            src_cfg = src_dir / "config.json"
            dst_cfg = dst_dir / "config.json"
            if src_cfg.exists():
                dst_cfg.write_text(src_cfg.read_text(encoding="utf-8"), encoding="utf-8")
        except Exception:
            pass

        if copy_db:
            try:
                src_db = src_dir / "pdm.db"
                dst_db = dst_dir / "pdm.db"
                if src_db.exists():
                    dst_db.write_bytes(src_db.read_bytes())
            except Exception:
                pass

        return ws

    def delete(self, ws_id: str, delete_folder: bool = False) -> None:
        ws = self.get(ws_id)
        if not ws:
            return
        self._index.pop(ws_id, None)
        self._save_index()

        if self.get_current_id() == ws_id:
            nxt = self.list()
            if nxt:
                self.set_current(nxt[0].id)
            else:
                self.current_path.write_text("", encoding="utf-8")

        if delete_folder:
            try:
                import shutil as _shutil
                _shutil.rmtree(self._resolve_ws_dir(ws), ignore_errors=True)
            except Exception:
                pass

    def workspace_dir(self, ws_id: str) -> Path:
        ws = self.get(ws_id)
        if not ws:
            raise ValueError("Workspace non trovata")
        ws_dir = self._resolve_ws_dir(ws)
        ws_dir.mkdir(parents=True, exist_ok=True)
        self._ensure_ws_subdirs(ws_dir)
        return ws_dir

    def config_path(self, ws_id: str) -> Path:
        return self.workspace_dir(ws_id) / "config.json"

    def db_path(self, ws_id: str) -> Path:
        return self.workspace_dir(ws_id) / "pdm.db"

    def backups_dir(self, ws_id: str) -> Path:
        p = self.workspace_dir(ws_id) / "backups"
        p.mkdir(parents=True, exist_ok=True)
        return p

    def meta_path(self, ws_id: str) -> Path:
        return self.workspace_dir(ws_id) / "workspace_meta.json"

    def read_meta(self, ws_id: str) -> Dict:
        p = self.meta_path(ws_id)
        if p.exists():
            try:
                return json.loads(p.read_text(encoding="utf-8"))
            except Exception:
                return {}
        return {}

    def write_meta(self, ws_id: str, meta: Dict) -> None:
        p = self.meta_path(ws_id)
        p.write_text(json.dumps(meta, indent=2, ensure_ascii=False), encoding="utf-8")
