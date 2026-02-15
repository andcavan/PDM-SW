from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Dict
from datetime import datetime, date
import json
import zipfile

from .version import __version__
from .workspace import WorkspaceManager
from .store import Store


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d_%H-%M-%S")


@dataclass
class BackupResult:
    ok: bool
    message: str
    path: str = ""


class BackupManager:
    def __init__(self, ws_mgr: WorkspaceManager, ws_id: str, store: Store, retention_total: int = 30):
        self.ws_mgr = ws_mgr
        self.ws_id = ws_id
        self.store = store
        self.retention_total = max(1, int(retention_total))

    def _meta(self) -> Dict:
        return self.ws_mgr.read_meta(self.ws_id)

    def _save_meta(self, meta: Dict) -> None:
        self.ws_mgr.write_meta(self.ws_id, meta)

    def backup_now(self, reason: str, force: bool = False) -> BackupResult:
        if not force and not self.store.dirty:
            return BackupResult(True, "Nessuna modifica: backup non necessario.")

        backups_dir = self.ws_mgr.backups_dir(self.ws_id)
        backups_dir.mkdir(parents=True, exist_ok=True)

        tmp_db = backups_dir / f"tmp_{_ts()}.db"
        zip_path = backups_dir / f"{_ts()}_{reason}.zip"

        try:
            self.store.backup_sqlite_to(tmp_db)
            cfg_path = self.ws_mgr.config_path(self.ws_id)
            manifest = {
                "workspace_id": self.ws_id,
                "reason": reason,
                "created_at": datetime.now().isoformat(timespec="seconds"),
                "app_version": __version__,
            }

            with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
                z.write(tmp_db, arcname="pdm.db")
                if cfg_path.exists():
                    z.write(cfg_path, arcname="config.json")
                z.writestr("manifest.json", json.dumps(manifest, indent=2, ensure_ascii=False))

            try:
                tmp_db.unlink(missing_ok=True)  # py3.11+
            except Exception:
                try:
                    tmp_db.unlink()
                except Exception:
                    pass

            self._enforce_retention(backups_dir)

            # reset dirty after successful backup
            self.store.clear_dirty()
            return BackupResult(True, "Backup creato.", str(zip_path))
        except Exception as e:
            try:
                tmp_db.unlink(missing_ok=True)
            except Exception:
                pass
            return BackupResult(False, f"Backup fallito: {e}")

    def maybe_daily_backup(self) -> Optional[BackupResult]:
        meta = self._meta()
        today = date.today().isoformat()
        last = str(meta.get("last_daily_backup", ""))
        if last == today:
            return None
        if not self.store.dirty:
            return None
        res = self.backup_now("daily", force=True)
        if res.ok:
            meta["last_daily_backup"] = today
            self._save_meta(meta)
        return res

    def _enforce_retention(self, backups_dir: Path) -> None:
        zips = sorted(backups_dir.glob("*.zip"), key=lambda p: p.name, reverse=True)
        for p in zips[self.retention_total:]:
            try:
                p.unlink()
            except Exception:
                pass
