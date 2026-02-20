from __future__ import annotations

import sqlite3
import json
from pathlib import Path
from typing import List, Optional, Tuple, Dict, Any
from datetime import datetime, timedelta

from .models import Document, DocType, State


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _now_plus(seconds: int) -> str:
    s = max(10, int(seconds or 0))
    return (datetime.now() + timedelta(seconds=s)).isoformat(timespec="seconds")


class Store:
    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(str(self.db_path), timeout=30.0)
        self.conn.row_factory = sqlite3.Row
        self.dirty = False
        self._init_db()

    def close(self) -> None:
        try:
            self.conn.close()
        except Exception:
            pass

    def _init_db(self) -> None:
        c = self.conn.cursor()
        c.execute("PRAGMA busy_timeout=30000;")
        try:
            c.execute("PRAGMA journal_mode=WAL;")
        except Exception:
            pass
        c.execute("""
        CREATE TABLE IF NOT EXISTS machines(
            mmm TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS groups(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            mmm TEXT NOT NULL,
            gggg TEXT NOT NULL,
            name TEXT NOT NULL,
            created_at TEXT NOT NULL,
            UNIQUE(mmm, gggg),
            FOREIGN KEY(mmm) REFERENCES machines(mmm) ON DELETE CASCADE
        );
        """)
        c.execute("""
        CREATE TABLE IF NOT EXISTS seq_counters(
            mmm TEXT NOT NULL,
            gggg TEXT NOT NULL,
            vvv TEXT NOT NULL DEFAULT '',
            next_part INTEGER NOT NULL,
            next_assy INTEGER NOT NULL,
            PRIMARY KEY(mmm, gggg, vvv)
        );
        """)
        # Migrazione seq_counters: da PK(mmm, gggg) a PK(mmm, gggg, vvv)
        try:
            cols = [r["name"] for r in self.conn.execute("PRAGMA table_info(seq_counters);").fetchall()]
            if "vvv" not in cols:
                c.execute("""
                CREATE TABLE IF NOT EXISTS seq_counters_new(
                    mmm TEXT NOT NULL,
                    gggg TEXT NOT NULL,
                    vvv TEXT NOT NULL DEFAULT '',
                    next_part INTEGER NOT NULL,
                    next_assy INTEGER NOT NULL,
                    PRIMARY KEY(mmm, gggg, vvv)
                );
                """)
                c.execute("INSERT INTO seq_counters_new(mmm, gggg, vvv, next_part, next_assy) SELECT mmm, gggg, '' as vvv, next_part, next_assy FROM seq_counters;")
                c.execute("DROP TABLE seq_counters;")
                c.execute("ALTER TABLE seq_counters_new RENAME TO seq_counters;")
        except Exception:
            pass

        c.execute("""
        CREATE TABLE IF NOT EXISTS ver_counters(
            mmm TEXT NOT NULL,
            gggg TEXT NOT NULL DEFAULT '',
            doc_type TEXT NOT NULL,
            next_ver INTEGER NOT NULL DEFAULT 1,
            PRIMARY KEY(mmm, gggg, doc_type)
        );
        """)

        c.execute("""
        CREATE TABLE IF NOT EXISTS documents(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT NOT NULL UNIQUE,
            doc_type TEXT NOT NULL,
            mmm TEXT NOT NULL,
            gggg TEXT NOT NULL,
            seq INTEGER NOT NULL,
            vvv TEXT NOT NULL DEFAULT '',
            revision INTEGER NOT NULL DEFAULT 0,
            state TEXT NOT NULL DEFAULT 'WIP',
            obs_prev_state TEXT NOT NULL DEFAULT '',
            description TEXT NOT NULL DEFAULT '',
            checked_out INTEGER NOT NULL DEFAULT 0,
            checkout_owner_user TEXT NOT NULL DEFAULT '',
            checkout_owner_host TEXT NOT NULL DEFAULT '',
            checkout_at TEXT NOT NULL DEFAULT '',

            file_wip_path TEXT NOT NULL DEFAULT '',
            file_rel_path TEXT NOT NULL DEFAULT '',
            file_inrev_path TEXT NOT NULL DEFAULT '',

            file_wip_drw_path TEXT NOT NULL DEFAULT '',
            file_rel_drw_path TEXT NOT NULL DEFAULT '',
            file_inrev_drw_path TEXT NOT NULL DEFAULT '',

            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        """)

        # Migrazione documents: aggiunta obs_prev_state
        try:
            cols = [r["name"] for r in self.conn.execute("PRAGMA table_info(documents);").fetchall()]
            if "obs_prev_state" not in cols:
                c.execute("ALTER TABLE documents ADD COLUMN obs_prev_state TEXT NOT NULL DEFAULT '';")
            if "checked_out" not in cols:
                c.execute("ALTER TABLE documents ADD COLUMN checked_out INTEGER NOT NULL DEFAULT 0;")
            if "checkout_owner_user" not in cols:
                c.execute("ALTER TABLE documents ADD COLUMN checkout_owner_user TEXT NOT NULL DEFAULT '';")
            if "checkout_owner_host" not in cols:
                c.execute("ALTER TABLE documents ADD COLUMN checkout_owner_host TEXT NOT NULL DEFAULT '';")
            if "checkout_at" not in cols:
                c.execute("ALTER TABLE documents ADD COLUMN checkout_at TEXT NOT NULL DEFAULT '';")
        except Exception:
            pass

        c.execute("""
        CREATE TABLE IF NOT EXISTS doc_custom_values(
            code TEXT NOT NULL,
            prop_name TEXT NOT NULL,
            value TEXT NOT NULL DEFAULT '',
            updated_at TEXT NOT NULL,
            PRIMARY KEY(code, prop_name)
        );
        """)
        c.execute("CREATE INDEX IF NOT EXISTS idx_doc_custom_values_prop ON doc_custom_values(prop_name);")

        c.execute("""
        CREATE TABLE IF NOT EXISTS document_state_notes(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT NOT NULL,
            created_at TEXT NOT NULL,
            event_type TEXT NOT NULL,
            from_state TEXT NOT NULL,
            to_state TEXT NOT NULL,
            note TEXT NOT NULL,
            rev_before INTEGER NOT NULL DEFAULT 0,
            rev_after INTEGER NOT NULL DEFAULT 0
        );
        """)
        c.execute("CREATE INDEX IF NOT EXISTS idx_document_state_notes_code_time ON document_state_notes(code, created_at DESC);")

        c.execute("""
        CREATE TABLE IF NOT EXISTS document_locks(
            code TEXT PRIMARY KEY,
            owner_session TEXT NOT NULL,
            owner_user TEXT NOT NULL DEFAULT '',
            owner_host TEXT NOT NULL DEFAULT '',
            acquired_at TEXT NOT NULL,
            updated_at TEXT NOT NULL,
            expires_at TEXT NOT NULL
        );
        """)
        c.execute("CREATE INDEX IF NOT EXISTS idx_document_locks_expires ON document_locks(expires_at);")

        c.execute("""
        CREATE TABLE IF NOT EXISTS activity_log(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            workspace_id TEXT NOT NULL DEFAULT '',
            session_id TEXT NOT NULL DEFAULT '',
            user_id TEXT NOT NULL DEFAULT '',
            user_display TEXT NOT NULL DEFAULT '',
            host TEXT NOT NULL DEFAULT '',
            action TEXT NOT NULL,
            code TEXT NOT NULL DEFAULT '',
            status TEXT NOT NULL DEFAULT 'OK',
            message TEXT NOT NULL DEFAULT '',
            details_json TEXT NOT NULL DEFAULT ''
        );
        """)
        c.execute("CREATE INDEX IF NOT EXISTS idx_activity_log_time ON activity_log(created_at DESC, id DESC);")
        c.execute("CREATE INDEX IF NOT EXISTS idx_activity_log_action ON activity_log(action);")
        c.execute("CREATE INDEX IF NOT EXISTS idx_activity_log_code ON activity_log(code);")

        self.conn.commit()

    def _mark_dirty(self) -> None:
        self.dirty = True

    def clear_dirty(self) -> None:
        self.dirty = False

    # --- backup
    def backup_sqlite_to(self, dest_db_path: Path) -> None:
        dest_db_path = Path(dest_db_path)
        dest_db_path.parent.mkdir(parents=True, exist_ok=True)
        # checkpoint WAL for consistency (best effort)
        try:
            self.conn.execute("PRAGMA wal_checkpoint(FULL);")
        except Exception:
            pass
        self.conn.commit()
        with sqlite3.connect(str(dest_db_path)) as dest:
            self.conn.backup(dest)

    # --- machines & groups
    def list_machines(self) -> List[Tuple[str, str]]:
        cur = self.conn.execute("SELECT mmm, name FROM machines ORDER BY mmm;")
        return [(r["mmm"], r["name"]) for r in cur.fetchall()]

    def add_machine(self, mmm: str, name: str) -> None:
        self.conn.execute(
            "INSERT OR REPLACE INTO machines(mmm, name, created_at) VALUES(?, ?, ?);",
            (mmm, name, _now()),
        )
        self.conn.commit()
        self._mark_dirty()

    def delete_machine(self, mmm: str) -> None:
        self.conn.execute("DELETE FROM machines WHERE mmm=?;", (mmm,))
        self.conn.commit()
        self._mark_dirty()

    def list_groups(self, mmm: str) -> List[Tuple[str, str]]:
        cur = self.conn.execute("SELECT gggg, name FROM groups WHERE mmm=? ORDER BY gggg;", (mmm,))
        return [(r["gggg"], r["name"]) for r in cur.fetchall()]

    def add_group(self, mmm: str, gggg: str, name: str) -> None:
        self.conn.execute(
            "INSERT OR REPLACE INTO groups(mmm, gggg, name, created_at) VALUES(?, ?, ?, ?);",
            (mmm, gggg, name, _now()),
        )
        self.conn.commit()
        self._mark_dirty()

    def delete_group(self, mmm: str, gggg: str) -> None:
        self.conn.execute("DELETE FROM groups WHERE mmm=? AND gggg=?;", (mmm, gggg))
        self.conn.commit()
        self._mark_dirty()

    # --- sequence allocation
    def allocate_seq(self, mmm: str, gggg: str, vvv: str, doc_type: DocType) -> int:
        dt = str(doc_type).upper()
        if dt in ("PRT", "PART", "SLDPRT"):
            dt = "PART"
        elif dt in ("ASM", "ASSY", "SLDASM"):
            dt = "ASSY"
        doc_type = dt
        cur = self.conn.cursor()
        try:
            cur.execute("BEGIN IMMEDIATE;")
            row = cur.execute(
                "SELECT next_part, next_assy FROM seq_counters WHERE mmm=? AND gggg=? AND vvv=?;",
                (mmm, gggg, vvv),
            ).fetchone()
            if row is None:
                next_part, next_assy = 1, 9999
                cur.execute(
                    "INSERT INTO seq_counters(mmm, gggg, vvv, next_part, next_assy) VALUES(?, ?, ?, ?, ?);",
                    (mmm, gggg, vvv, next_part, next_assy),
                )
            else:
                next_part, next_assy = int(row["next_part"]), int(row["next_assy"])

            if doc_type == "PART":
                if next_part > 9999:
                    raise ValueError("Sequenza PART esaurita (oltre 9999)")
                seq = next_part
                next_part += 1
            else:
                if next_assy < 1:
                    raise ValueError("Sequenza ASSY esaurita (sotto 0001)")
                seq = next_assy
                next_assy -= 1

            cur.execute(
                "UPDATE seq_counters SET next_part=?, next_assy=? WHERE mmm=? AND gggg=? AND vvv=?;",
                (next_part, next_assy, mmm, gggg, vvv),
            )
            self.conn.commit()
        except Exception:
            try:
                self.conn.rollback()
            except Exception:
                pass
            raise
        self._mark_dirty()
        return seq


    def peek_seq(self, mmm: str, gggg: str, vvv: str, doc_type: DocType) -> int:
        dt = str(doc_type).upper()
        if dt in ("PRT", "PART", "SLDPRT"):
            dt = "PART"
        elif dt in ("ASM", "ASSY", "SLDASM"):
            dt = "ASSY"
        doc_type = dt
        """Ritorna il prossimo progressivo SENZA incrementare il contatore."""
        row = self.conn.execute(
            "SELECT next_part, next_assy FROM seq_counters WHERE mmm=? AND gggg=? AND vvv=?;",
            (mmm, gggg, vvv),
        ).fetchone()
        if row is None:
            next_part, next_assy = 1, 9999
        else:
            next_part, next_assy = int(row["next_part"]), int(row["next_assy"])

        if doc_type == "PART":
            if next_part > 9999:
                raise ValueError("Sequenza PART esaurita (oltre 9999)")
            return next_part
        else:
            if next_assy < 1:
                raise ValueError("Sequenza ASSY esaurita (sotto 0001)")
            return next_assy

    def allocate_ver_seq(self, mmm: str, gggg: str, doc_type: DocType) -> int:
        """Alloca progressivo versione per MACHINE o GROUP."""
        dt = str(doc_type).upper()
        if dt not in ("MACHINE", "GROUP"):
            raise ValueError(f"allocate_ver_seq valido solo per MACHINE/GROUP, ricevuto: {dt}")
        
        # MACHINE: gggg vuoto, GROUP: gggg valorizzato
        gggg_key = "" if dt == "MACHINE" else gggg
        
        cur = self.conn.cursor()
        try:
            cur.execute("BEGIN IMMEDIATE;")
            row = cur.execute(
                "SELECT next_ver FROM ver_counters WHERE mmm=? AND gggg=? AND doc_type=?;",
                (mmm, gggg_key, dt),
            ).fetchone()
            if row is None:
                next_ver = 1
                cur.execute(
                    "INSERT INTO ver_counters(mmm, gggg, doc_type, next_ver) VALUES(?, ?, ?, ?);",
                    (mmm, gggg_key, dt, next_ver + 1),
                )
            else:
                next_ver = int(row["next_ver"])
                cur.execute(
                    "UPDATE ver_counters SET next_ver=? WHERE mmm=? AND gggg=? AND doc_type=?;",
                    (next_ver + 1, mmm, gggg_key, dt),
                )
            self.conn.commit()
        except Exception:
            try:
                self.conn.rollback()
            except Exception:
                pass
            raise
        self._mark_dirty()
        return next_ver

    # --- documents
    def add_document(self, doc: Document) -> int:
        now = _now()
        cur = self.conn.execute(
            """
            INSERT INTO documents(
                code, doc_type, mmm, gggg, seq, vvv, revision, state, description,
                checked_out, checkout_owner_user, checkout_owner_host, checkout_at,
                file_wip_path, file_rel_path, file_inrev_path,
                file_wip_drw_path, file_rel_drw_path, file_inrev_drw_path,
                created_at, updated_at
            ) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """,
            (
                doc.code, doc.doc_type, doc.mmm, doc.gggg, doc.seq, doc.vvv, doc.revision, doc.state, doc.description,
                (1 if bool(getattr(doc, "checked_out", False)) else 0),
                str(getattr(doc, "checkout_owner_user", "") or ""),
                str(getattr(doc, "checkout_owner_host", "") or ""),
                str(getattr(doc, "checkout_at", "") or ""),
                doc.file_wip_path, doc.file_rel_path, doc.file_inrev_path,
                doc.file_wip_drw_path, doc.file_rel_drw_path, doc.file_inrev_drw_path,
                now, now
            )
        )
        self.conn.commit()
        self._mark_dirty()
        return int(cur.lastrowid)

    def get_document(self, code: str) -> Optional[Document]:
        r = self.conn.execute("SELECT * FROM documents WHERE code=?;", (code,)).fetchone()
        return self._row_to_doc(r) if r else None

    def update_document(self, code: str, **fields) -> None:
        if not fields:
            return
        fields["updated_at"] = _now()
        keys = list(fields.keys())
        sql = "UPDATE documents SET " + ", ".join([f"{k}=?" for k in keys]) + " WHERE code=?;"
        vals = [fields[k] for k in keys] + [code]
        self.conn.execute(sql, vals)
        self.conn.commit()
        self._mark_dirty()

    def checkout_document(self, code: str, owner_user: str, owner_host: str) -> Tuple[bool, str, Dict[str, str]]:
        code_u = (code or "").strip()
        usr = (owner_user or "").strip()
        host = (owner_host or "").strip()
        if not code_u:
            return False, "Codice mancante.", {}

        now = _now()
        cur = self.conn.cursor()
        try:
            cur.execute("BEGIN IMMEDIATE;")
            row = cur.execute(
                """
                SELECT code, state, checked_out, checkout_owner_user, checkout_owner_host, checkout_at
                FROM documents
                WHERE code=?;
                """,
                (code_u,),
            ).fetchone()
            if row is None:
                self.conn.commit()
                return False, "Documento non trovato.", {}

            state = str(row["state"] or "").strip().upper()
            if state in ("REL", "OBS"):
                self.conn.commit()
                return False, f"Checkout non consentito su stato {state}.", {}

            already = int(row["checked_out"] or 0) != 0
            holder = {
                "code": str(row["code"] or ""),
                "checkout_owner_user": str(row["checkout_owner_user"] or ""),
                "checkout_owner_host": str(row["checkout_owner_host"] or ""),
                "checkout_at": str(row["checkout_at"] or ""),
            }
            if already:
                holder_user = holder["checkout_owner_user"]
                if holder_user and usr and holder_user == usr:
                    cur.execute(
                        """
                        UPDATE documents
                        SET checked_out=1, checkout_owner_user=?, checkout_owner_host=?, checkout_at=?, updated_at=?
                        WHERE code=?;
                        """,
                        (usr, host, now, now, code_u),
                    )
                    self.conn.commit()
                    self._mark_dirty()
                    return True, "CHECKOUT_REFRESHED", holder
                self.conn.commit()
                return False, "CHECKOUT_BY_OTHER", holder

            cur.execute(
                """
                UPDATE documents
                SET checked_out=1, checkout_owner_user=?, checkout_owner_host=?, checkout_at=?, updated_at=?
                WHERE code=?;
                """,
                (usr, host, now, now, code_u),
            )
            self.conn.commit()
            self._mark_dirty()
            return True, "CHECKOUT_OK", {}
        except Exception as e:
            try:
                self.conn.rollback()
            except Exception:
                pass
            return False, f"CHECKOUT_ERROR: {e}", {}

    def checkin_document(self, code: str, owner_user: str, force: bool = False) -> Tuple[bool, str, Dict[str, str]]:
        code_u = (code or "").strip()
        usr = (owner_user or "").strip()
        if not code_u:
            return False, "Codice mancante.", {}

        now = _now()
        cur = self.conn.cursor()
        try:
            cur.execute("BEGIN IMMEDIATE;")
            row = cur.execute(
                """
                SELECT code, checked_out, checkout_owner_user, checkout_owner_host, checkout_at
                FROM documents
                WHERE code=?;
                """,
                (code_u,),
            ).fetchone()
            if row is None:
                self.conn.commit()
                return False, "Documento non trovato.", {}

            already = int(row["checked_out"] or 0) != 0
            if not already:
                self.conn.commit()
                return True, "CHECKIN_ALREADY", {}

            holder = {
                "code": str(row["code"] or ""),
                "checkout_owner_user": str(row["checkout_owner_user"] or ""),
                "checkout_owner_host": str(row["checkout_owner_host"] or ""),
                "checkout_at": str(row["checkout_at"] or ""),
            }
            holder_user = holder["checkout_owner_user"]
            if (not force) and holder_user and usr and holder_user != usr:
                self.conn.commit()
                return False, "CHECKOUT_BY_OTHER", holder
            if (not force) and holder_user and (not usr):
                self.conn.commit()
                return False, "CHECKIN_OWNER_REQUIRED", holder

            cur.execute(
                """
                UPDATE documents
                SET checked_out=0, checkout_owner_user='', checkout_owner_host='', checkout_at='', updated_at=?
                WHERE code=?;
                """,
                (now, code_u),
            )
            self.conn.commit()
            self._mark_dirty()
            return True, "CHECKIN_OK", holder
        except Exception as e:
            try:
                self.conn.rollback()
            except Exception:
                pass
            return False, f"CHECKIN_ERROR: {e}", {}

    def clear_document_checkout(self, code: str) -> bool:
        code_u = (code or "").strip()
        if not code_u:
            return False
        cur = self.conn.cursor()
        cur.execute(
            """
            UPDATE documents
            SET checked_out=0, checkout_owner_user='', checkout_owner_host='', checkout_at='', updated_at=?
            WHERE code=?;
            """,
            (_now(), code_u),
        )
        self.conn.commit()
        if cur.rowcount:
            self._mark_dirty()
        return bool(cur.rowcount)

    def list_documents(self, include_obs: bool = False) -> List[Document]:
        if include_obs:
            cur = self.conn.execute("SELECT * FROM documents ORDER BY updated_at DESC;")
        else:
            cur = self.conn.execute("SELECT * FROM documents WHERE state != 'OBS' ORDER BY updated_at DESC;")
        return [self._row_to_doc(r) for r in cur.fetchall()]

    def search_documents(
        self,
        query: str = "",
        mmm: str = "",
        gggg: str = "",
        vvv: str = "",
        state: str = "",
        doc_type: str = "",
        include_obs: bool = False,
        **kwargs,
    ) -> List[Document]:
        """Ricerca documenti.

        Backward compatible:
        - accetta alias 'text' in kwargs come sinonimo di query.
        - parametri extra ignorati.
        """
        if not query and "text" in kwargs:
            try:
                query = str(kwargs.get("text") or "")
            except Exception:
                query = ""

        where = []
        params: List[Any] = []
        if query:
            where.append("(code LIKE ? OR description LIKE ?)")
            params.extend([f"%{query}%", f"%{query}%"])
        if mmm:
            where.append("mmm=?"); params.append(mmm)
        if gggg:
            where.append("gggg=?"); params.append(gggg)
        if vvv:
            where.append("vvv=?"); params.append(vvv)
        if state:
            where.append("state=?"); params.append(state)
        if doc_type:
            where.append("doc_type=?"); params.append(doc_type)
        if not include_obs:
            where.append("state!='OBS'")
        sql = "SELECT * FROM documents"
        if where:
            sql += " WHERE " + " AND ".join(where)
        sql += " ORDER BY updated_at DESC;"
        cur = self.conn.execute(sql, params)
        return [self._row_to_doc(r) for r in cur.fetchall()]


    def _row_to_doc(self, r: sqlite3.Row) -> Document:
        return Document(
            id=int(r["id"]),
            code=str(r["code"]),
            doc_type=str(r["doc_type"]),
            mmm=str(r["mmm"]),
            gggg=str(r["gggg"]),
            seq=int(r["seq"]),
            vvv=str(r["vvv"] or ""),
            revision=int(r["revision"]),
            state=str(r["state"]),
            obs_prev_state=str(r["obs_prev_state"] if "obs_prev_state" in r.keys() else ""),
            description=str(r["description"] or ""),
            checked_out=(int(r["checked_out"] or 0) != 0) if "checked_out" in r.keys() else False,
            checkout_owner_user=str(r["checkout_owner_user"] if "checkout_owner_user" in r.keys() else ""),
            checkout_owner_host=str(r["checkout_owner_host"] if "checkout_owner_host" in r.keys() else ""),
            checkout_at=str(r["checkout_at"] if "checkout_at" in r.keys() else ""),

            file_wip_path=str(r["file_wip_path"] or ""),
            file_rel_path=str(r["file_rel_path"] or ""),
            file_inrev_path=str(r["file_inrev_path"] or ""),

            file_wip_drw_path=str(r["file_wip_drw_path"] or ""),
            file_rel_drw_path=str(r["file_rel_drw_path"] or ""),
            file_inrev_drw_path=str(r["file_inrev_drw_path"] or ""),

            created_at=str(r["created_at"]),
            updated_at=str(r["updated_at"]),
        )


    # ---- Custom properties values per documento ----
    def set_custom_value(self, code: str, prop_name: str, value: str) -> None:
        code = (code or "").strip()
        prop = (prop_name or "").strip().upper()
        if not code or not prop:
            return
        v = "" if value is None else str(value)
        c = self.conn.cursor()
        c.execute(
            "INSERT INTO doc_custom_values(code, prop_name, value, updated_at) VALUES(?,?,?,?) "
            "ON CONFLICT(code, prop_name) DO UPDATE SET value=excluded.value, updated_at=excluded.updated_at;",
            (code, prop, v, _now()),
        )
        self.conn.commit()
        self._mark_dirty()

    def get_custom_value(self, code: str, prop_name: str) -> str:
        code = (code or "").strip()
        prop = (prop_name or "").strip().upper()
        if not code or not prop:
            return ""
        r = self.conn.execute(
            "SELECT value FROM doc_custom_values WHERE code=? AND prop_name=?;",
            (code, prop),
        ).fetchone()
        return str(r["value"]) if r else ""

    def get_custom_values(self, code: str) -> Dict[str, str]:
        code = (code or "").strip()
        if not code:
            return {}
        rows = self.conn.execute(
            "SELECT prop_name, value FROM doc_custom_values WHERE code=?;",
            (code,),
        ).fetchall()
        return {str(r["prop_name"]): str(r["value"]) for r in rows}


    def get_custom_values_bulk(self, codes: List[str], prop_names: List[str]) -> Dict[str, Dict[str, str]]:
        """Recupera valori custom (doc_custom_values) per più codici e proprietà (ottimizzato).

        Ritorna: {code: {PROP: value}} dove PROP è uppercase.
        """
        codes_u = [str(c).strip() for c in (codes or []) if str(c).strip()]
        props_u = [str(p).strip().upper() for p in (prop_names or []) if str(p).strip()]
        if not codes_u:
            return {}
        if not props_u:
            return {c: {} for c in codes_u}

        out: Dict[str, Dict[str, str]] = {c: {} for c in codes_u}

        def chunks(lst, n):
            for i in range(0, len(lst), n):
                yield lst[i:i+n]

        for ch in chunks(codes_u, 400):
            ph_codes = ",".join(["?"] * len(ch))
            ph_props = ",".join(["?"] * len(props_u))
            sql = f"SELECT code, prop_name, value FROM doc_custom_values WHERE code IN ({ph_codes}) AND prop_name IN ({ph_props});"
            params = tuple(ch) + tuple(props_u)
            rows = self.conn.execute(sql, params).fetchall()
            for r in rows:
                c = str(r["code"])
                pn = str(r["prop_name"]).strip().upper()
                v = "" if r["value"] is None else str(r["value"])
                if c in out:
                    out[c][pn] = v
        return out

    def delete_custom_property_values(self, prop_name: str) -> None:
        prop = (prop_name or "").strip().upper()
        if not prop:
            return
        c = self.conn.cursor()
        c.execute("DELETE FROM doc_custom_values WHERE prop_name=?;", (prop,))
        self.conn.commit()
        self._mark_dirty()

    # ---- Workflow state notes ----
    def add_state_note(
        self,
        code: str,
        event_type: str,
        from_state: str,
        to_state: str,
        note: str,
        rev_before: int,
        rev_after: int,
    ) -> int:
        code_u = (code or "").strip()
        if not code_u:
            raise ValueError("Code mancante per nota stato.")
        note_u = (note or "").strip()
        if not note_u:
            raise ValueError("Nota stato vuota.")

        ev = (event_type or "").strip().upper()
        f_state = (from_state or "").strip().upper()
        t_state = (to_state or "").strip().upper()

        cur = self.conn.execute(
            """
            INSERT INTO document_state_notes(
                code, created_at, event_type, from_state, to_state, note, rev_before, rev_after
            ) VALUES(?, ?, ?, ?, ?, ?, ?, ?);
            """,
            (
                code_u,
                _now(),
                ev,
                f_state,
                t_state,
                note_u,
                int(rev_before),
                int(rev_after),
            ),
        )
        self.conn.commit()
        self._mark_dirty()
        return int(cur.lastrowid)

    def list_state_notes(self, code: str, limit: int = 200) -> List[Dict[str, Any]]:
        code_u = (code or "").strip()
        if not code_u:
            return []
        lim = max(1, int(limit or 200))
        rows = self.conn.execute(
            """
            SELECT id, code, created_at, event_type, from_state, to_state, note, rev_before, rev_after
            FROM document_state_notes
            WHERE code=?
            ORDER BY created_at DESC, id DESC
            LIMIT ?;
            """,
            (code_u, lim),
        ).fetchall()
        out: List[Dict[str, Any]] = []
        for r in rows:
            out.append(
                {
                    "id": int(r["id"]),
                    "code": str(r["code"]),
                    "created_at": str(r["created_at"]),
                    "event_type": str(r["event_type"]),
                    "from_state": str(r["from_state"]),
                    "to_state": str(r["to_state"]),
                    "note": str(r["note"]),
                    "rev_before": int(r["rev_before"]),
                    "rev_after": int(r["rev_after"]),
                }
            )
        return out

    # ---- Documento lock condiviso (multi-utente) ----
    def acquire_document_lock(
        self,
        code: str,
        owner_session: str,
        owner_user: str,
        owner_host: str,
        ttl_seconds: int = 1200,
    ) -> Tuple[bool, str, Dict[str, str]]:
        code_u = (code or "").strip()
        sess = (owner_session or "").strip()
        usr = (owner_user or "").strip()
        host = (owner_host or "").strip()
        if not code_u:
            return False, "Codice lock mancante.", {}
        if not sess:
            return False, "Sessione lock mancante.", {}

        now = _now()
        expires = _now_plus(ttl_seconds)
        cur = self.conn.cursor()
        try:
            cur.execute("BEGIN IMMEDIATE;")
            cur.execute("DELETE FROM document_locks WHERE expires_at <= ?;", (now,))
            row = cur.execute(
                """
                SELECT code, owner_session, owner_user, owner_host, acquired_at, updated_at, expires_at
                FROM document_locks
                WHERE code=?;
                """,
                (code_u,),
            ).fetchone()

            if row is None:
                cur.execute(
                    """
                    INSERT INTO document_locks(code, owner_session, owner_user, owner_host, acquired_at, updated_at, expires_at)
                    VALUES(?, ?, ?, ?, ?, ?, ?);
                    """,
                    (code_u, sess, usr, host, now, now, expires),
                )
                self.conn.commit()
                return True, "LOCK_ACQUIRED", {}

            holder = {
                "code": str(row["code"]),
                "owner_session": str(row["owner_session"]),
                "owner_user": str(row["owner_user"]),
                "owner_host": str(row["owner_host"]),
                "acquired_at": str(row["acquired_at"]),
                "updated_at": str(row["updated_at"]),
                "expires_at": str(row["expires_at"]),
            }

            if holder["owner_session"] == sess:
                cur.execute(
                    "UPDATE document_locks SET owner_user=?, owner_host=?, updated_at=?, expires_at=? WHERE code=?;",
                    (usr, host, now, expires, code_u),
                )
                self.conn.commit()
                return True, "LOCK_REFRESHED", holder

            self.conn.commit()
            return False, "LOCKED_BY_OTHER", holder
        except Exception as e:
            try:
                self.conn.rollback()
            except Exception:
                pass
            return False, f"LOCK_ERROR: {e}", {}

    def release_document_lock(self, code: str, owner_session: str) -> bool:
        code_u = (code or "").strip()
        sess = (owner_session or "").strip()
        if not code_u or not sess:
            return False
        cur = self.conn.cursor()
        cur.execute("DELETE FROM document_locks WHERE code=? AND owner_session=?;", (code_u, sess))
        self.conn.commit()
        return cur.rowcount > 0

    def release_session_locks(self, owner_session: str) -> int:
        sess = (owner_session or "").strip()
        if not sess:
            return 0
        cur = self.conn.cursor()
        cur.execute("DELETE FROM document_locks WHERE owner_session=?;", (sess,))
        self.conn.commit()
        return int(cur.rowcount or 0)

    def list_active_locks(self, limit: int = 500) -> List[Dict[str, str]]:
        lim = max(1, int(limit or 500))
        now = _now()
        try:
            self.conn.execute("DELETE FROM document_locks WHERE expires_at <= ?;", (now,))
            self.conn.commit()
        except Exception:
            pass
        rows = self.conn.execute(
            """
            SELECT code, owner_session, owner_user, owner_host, acquired_at, updated_at, expires_at
            FROM document_locks
            ORDER BY updated_at DESC, code ASC
            LIMIT ?;
            """,
            (lim,),
        ).fetchall()
        out: List[Dict[str, str]] = []
        for r in rows:
            out.append(
                {
                    "code": str(r["code"]),
                    "owner_session": str(r["owner_session"]),
                    "owner_user": str(r["owner_user"]),
                    "owner_host": str(r["owner_host"]),
                    "acquired_at": str(r["acquired_at"]),
                    "updated_at": str(r["updated_at"]),
                    "expires_at": str(r["expires_at"]),
                }
            )
        return out

    # ---- Activity log (audit operazioni) ----
    def add_activity(
        self,
        workspace_id: str,
        session_id: str,
        user_id: str,
        user_display: str,
        host: str,
        action: str,
        code: str = "",
        status: str = "OK",
        message: str = "",
        details: Optional[Dict[str, Any]] = None,
    ) -> int:
        act = (action or "").strip().upper()
        if not act:
            raise ValueError("Azione activity mancante.")
        details_json = ""
        if details:
            try:
                details_json = json.dumps(details, ensure_ascii=False, separators=(",", ":"))
            except Exception:
                details_json = ""
        cur = self.conn.execute(
            """
            INSERT INTO activity_log(
                created_at, workspace_id, session_id, user_id, user_display, host, action, code, status, message, details_json
            ) VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """,
            (
                _now(),
                str(workspace_id or ""),
                str(session_id or ""),
                str(user_id or ""),
                str(user_display or ""),
                str(host or ""),
                act,
                str(code or ""),
                str(status or "OK"),
                str(message or ""),
                details_json,
            ),
        )
        self.conn.commit()
        return int(cur.lastrowid)

    def list_recent_activity(self, limit: int = 200) -> List[Dict[str, Any]]:
        lim = max(1, int(limit or 200))
        rows = self.conn.execute(
            """
            SELECT id, created_at, workspace_id, session_id, user_id, user_display, host, action, code, status, message, details_json
            FROM activity_log
            ORDER BY created_at DESC, id DESC
            LIMIT ?;
            """,
            (lim,),
        ).fetchall()
        out: List[Dict[str, Any]] = []
        for r in rows:
            details = {}
            dj = str(r["details_json"] or "").strip()
            if dj:
                try:
                    details = json.loads(dj)
                except Exception:
                    details = {"raw": dj}
            out.append(
                {
                    "id": int(r["id"]),
                    "created_at": str(r["created_at"]),
                    "workspace_id": str(r["workspace_id"]),
                    "session_id": str(r["session_id"]),
                    "user_id": str(r["user_id"]),
                    "user_display": str(r["user_display"]),
                    "host": str(r["host"]),
                    "action": str(r["action"]),
                    "code": str(r["code"]),
                    "status": str(r["status"]),
                    "message": str(r["message"]),
                    "details": details,
                }
            )
        return out
