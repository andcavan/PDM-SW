from __future__ import annotations

import csv
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from tkinter import messagebox

from pdm_sw.models import Document


def warn(msg: str) -> None:
    messagebox.showwarning("PDM", msg)


def info(msg: str) -> None:
    messagebox.showinfo("PDM", msg)


class ReportMixin:
    def _report_dir(self) -> Path:
        p = self.ws_mgr.workspace_dir(self.ws_id) / "REPORTS"
        p.mkdir(parents=True, exist_ok=True)
        return p

    def _workflow_log_path(self) -> Path:
        p = self.ws_mgr.workspace_dir(self.ws_id) / "LOGS"
        p.mkdir(parents=True, exist_ok=True)
        return p / "workflow.log"

    def _workflow_log_line(self, msg: str) -> None:
        try:
            p = self._workflow_log_path()
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with p.open("a", encoding="utf-8") as f:
                f.write(f"{ts} | {msg.rstrip()}\n")
        except Exception:
            pass

    def _report_ts(self) -> str:
        return datetime.now().strftime("%Y%m%d_%H%M%S")

    def _report_token(self, value: str) -> str:
        v = (value or "").strip()
        if not v:
            return "NA"
        tok = "".join(ch if (ch.isalnum() or ch in ("-", "_")) else "_" for ch in v)
        while "__" in tok:
            tok = tok.replace("__", "_")
        tok = tok.strip("_")
        return tok[:120] if tok else "NA"

    def _generate_code_report(self):
        doc = self._load_selected_doc()
        if not doc:
            warn("Seleziona o carica un codice nel Workflow.")
            return

        report_dir = self._report_dir()
        ts = self._report_ts()
        base = f"report_codice_{self._report_token(doc.code)}_{ts}"

        txt_path = report_dir / f"{base}.txt"
        csv_path = report_dir / f"{base}.csv"
        notes_csv_path = report_dir / f"{base}_note.csv"
        custom_csv_path = report_dir / f"{base}_custom.csv"

        model_best, drw_best = self._best_model_and_drw_paths(doc)
        m_ok, d_ok = self._model_and_drawing_flags(doc)
        m_exists = "YES" if m_ok else "NO"
        d_exists = "YES" if d_ok else "NO"

        def _exists(path_s: str) -> str:
            return "YES" if (path_s and Path(path_s).is_file()) else "NO"

        def _shown_path(path_s: str) -> str:
            return str(path_s or "") if (path_s and Path(path_s).is_file()) else ""

        try:
            custom_vals = self.store.get_custom_values(doc.code) or {}
        except Exception:
            custom_vals = {}
        try:
            notes = self.store.list_state_notes(doc.code, limit=5000)
        except Exception:
            notes = []

        lines: list[str] = []
        lines.append(f"REPORT CODICE: {doc.code}")
        lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append("")
        lines.append("[ANAGRAFICA]")
        lines.append(f"Code: {doc.code}")
        lines.append(f"Doc type: {doc.doc_type}")
        lines.append(f"State: {doc.state}")
        lines.append(f"Revision: {int(doc.revision):02d}")
        lines.append(f"MMM/GGGG/VVV: {doc.mmm}/{doc.gggg}/{doc.vvv}")
        lines.append(f"Seq: {doc.seq:04d}")
        lines.append(f"Description: {doc.description}")
        lines.append(f"OBS previous state: {getattr(doc, 'obs_prev_state', '') or ''}")
        lines.append(f"Created at: {doc.created_at}")
        lines.append(f"Updated at: {doc.updated_at}")
        lines.append("")
        lines.append("[FILE]")
        lines.append(f"Best model path: {model_best}")
        lines.append(f"Best drawing path: {drw_best}")
        lines.append(f"Model exists (M): {m_exists}")
        lines.append(f"Drawing exists (D): {d_exists}")
        lines.append(f"MODEL WIP [{_exists(doc.file_wip_path)}]: {_shown_path(doc.file_wip_path)}")
        lines.append(f"MODEL REL [{_exists(doc.file_rel_path)}]: {_shown_path(doc.file_rel_path)}")
        lines.append(f"MODEL INREV [{_exists(doc.file_inrev_path)}]: {_shown_path(doc.file_inrev_path)}")
        lines.append(f"DRW WIP [{_exists(doc.file_wip_drw_path)}]: {_shown_path(doc.file_wip_drw_path)}")
        lines.append(f"DRW REL [{_exists(doc.file_rel_drw_path)}]: {_shown_path(doc.file_rel_drw_path)}")
        lines.append(f"DRW INREV [{_exists(doc.file_inrev_drw_path)}]: {_shown_path(doc.file_inrev_drw_path)}")
        lines.append("")
        lines.append("[CUSTOM VALUES]")
        if custom_vals:
            for k in sorted(custom_vals.keys()):
                lines.append(f"{k} = {custom_vals.get(k, '')}")
        else:
            lines.append("(none)")
        lines.append("")
        lines.append("[STATE NOTES]")
        if notes:
            for n in notes:
                ts_note = str(n.get("created_at", "")).replace("T", " ")
                lines.append(
                    f"{ts_note} | {n.get('event_type', '')} | "
                    f"{n.get('from_state', '')}->{n.get('to_state', '')} | "
                    f"REV {int(n.get('rev_before', 0)):02d}->{int(n.get('rev_after', 0)):02d}"
                )
                lines.append(str(n.get("note", "")).strip())
                lines.append("")
        else:
            lines.append("(none)")

        txt_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")

        with csv_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=[
                    "code",
                    "doc_type",
                    "state",
                    "revision",
                    "mmm",
                    "gggg",
                    "vvv",
                    "seq",
                    "description",
                    "obs_prev_state",
                    "created_at",
                    "updated_at",
                    "model_exists_m",
                    "drawing_exists_d",
                    "best_model_path",
                    "best_drawing_path",
                    "file_wip_path",
                    "file_rel_path",
                    "file_inrev_path",
                    "file_wip_drw_path",
                    "file_rel_drw_path",
                    "file_inrev_drw_path",
                ],
            )
            writer.writeheader()
            writer.writerow(
                {
                    "code": doc.code,
                    "doc_type": doc.doc_type,
                    "state": doc.state,
                    "revision": int(doc.revision),
                    "mmm": doc.mmm,
                    "gggg": doc.gggg,
                    "vvv": doc.vvv,
                    "seq": int(doc.seq),
                    "description": doc.description,
                    "obs_prev_state": getattr(doc, "obs_prev_state", "") or "",
                    "created_at": doc.created_at,
                    "updated_at": doc.updated_at,
                    "model_exists_m": m_exists,
                    "drawing_exists_d": d_exists,
                    "best_model_path": model_best,
                    "best_drawing_path": drw_best,
                    "file_wip_path": doc.file_wip_path,
                    "file_rel_path": doc.file_rel_path,
                    "file_inrev_path": doc.file_inrev_path,
                    "file_wip_drw_path": doc.file_wip_drw_path,
                    "file_rel_drw_path": doc.file_rel_drw_path,
                    "file_inrev_drw_path": doc.file_inrev_drw_path,
                }
            )

        with notes_csv_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=["created_at", "event_type", "from_state", "to_state", "rev_before", "rev_after", "note"],
            )
            writer.writeheader()
            for n in notes:
                writer.writerow(
                    {
                        "created_at": n.get("created_at", ""),
                        "event_type": n.get("event_type", ""),
                        "from_state": n.get("from_state", ""),
                        "to_state": n.get("to_state", ""),
                        "rev_before": n.get("rev_before", 0),
                        "rev_after": n.get("rev_after", 0),
                        "note": n.get("note", ""),
                    }
                )

        with custom_csv_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["prop_name", "value"])
            writer.writeheader()
            for k in sorted(custom_vals.keys()):
                writer.writerow({"prop_name": k, "value": custom_vals.get(k, "")})

        info(
            "Report codice generato.\n\n"
            f"TXT: {txt_path}\n"
            f"CSV: {csv_path}\n"
            f"NOTE: {notes_csv_path}\n"
            f"CUSTOM: {custom_csv_path}"
        )

    def _generate_hierarchy_report(self):
        include_obs = bool(self.hierarchy_include_obs_var.get()) if hasattr(self, "hierarchy_include_obs_var") else False

        report_dir = self._report_dir()
        ts = self._report_ts()
        base = f"report_generale_gerarchico_{ts}"
        txt_path = report_dir / f"{base}.txt"
        csv_path = report_dir / f"{base}.csv"

        machines_raw = self.store.list_machines()
        machine_names: dict[str, str] = {mmm: name for mmm, name in machines_raw}
        groups_by_machine: dict[str, dict[str, str]] = {}
        for mmm, _ in machines_raw:
            groups_by_machine[mmm] = {gggg: g_name for gggg, g_name in self.store.list_groups(mmm)}

        docs = self.store.list_documents(include_obs=include_obs)
        docs_by_pair: dict[tuple[str, str], list[Document]] = defaultdict(list)
        docs_machine: dict[str, list[Document]] = defaultdict(list)
        for d in docs:
            if d.doc_type == "MACHINE":
                docs_machine[d.mmm].append(d)
            else:
                docs_by_pair[(d.mmm, d.gggg)].append(d)
            machine_names.setdefault(d.mmm, "")
            groups_by_machine.setdefault(d.mmm, {})
            if d.gggg:
                groups_by_machine[d.mmm].setdefault(d.gggg, "")

        for pair in list(docs_by_pair.keys()):
            docs_by_pair[pair].sort(key=lambda d: (0 if d.doc_type == "PART" else (1 if d.doc_type == "ASSY" else 2), d.code))

        lines: list[str] = []
        lines.append("REPORT GENERALE GERARCHICO")
        lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"Include OBS: {include_obs}")
        lines.append(f"Total docs: {len(docs)}")
        lines.append("")

        if not machine_names:
            lines.append("(nessun MMM disponibile)")
        else:
            for mmm in sorted(machine_names.keys()):
                m_name = (machine_names.get(mmm) or "").strip()
                lines.append(f"MMM: {mmm} | Name: {m_name}")
                
                # MACHINE versions
                machine_docs = docs_machine.get(mmm, [])
                if machine_docs:
                    machine_docs.sort(key=lambda d: d.code)
                    for d in machine_docs:
                        m_ok, d_ok = self._model_and_drawing_flags(d)
                        m_val = "YES" if m_ok else "NO"
                        d_val = "YES" if d_ok else "NO"
                        lines.append(
                            f"  [MACHINE] {d.code} | {d.state} | REV {int(d.revision):02d} | "
                            f"M:{m_val} D:{d_val} | {d.description}"
                        )
                
                group_map = groups_by_machine.get(mmm, {})
                gggg_keys = sorted(group_map.keys())
                if not gggg_keys:
                    if not machine_docs:
                        lines.append("  (nessun GGGG o versione macchina)")
                    lines.append("")
                    continue
                for gggg in gggg_keys:
                    g_name = (group_map.get(gggg) or "").strip()
                    dlist = docs_by_pair.get((mmm, gggg), [])
                    part_count = sum(1 for d in dlist if d.doc_type == "PART")
                    assy_count = sum(1 for d in dlist if d.doc_type == "ASSY")
                    group_count = sum(1 for d in dlist if d.doc_type == "GROUP")
                    lines.append(f"  GGGG: {gggg} | Name: {g_name} | GROUP:{group_count} PART:{part_count} ASSY:{assy_count}")
                    if not dlist:
                        lines.append("    (nessun codice)")
                    for d in dlist:
                        m_ok, d_ok = self._model_and_drawing_flags(d)
                        m_val = "YES" if m_ok else "NO"
                        d_val = "YES" if d_ok else "NO"
                        prefix = "[GROUP] " if d.doc_type == "GROUP" else "    "
                        lines.append(
                            f"{prefix}{d.code} | {d.doc_type} | {d.state} | REV {int(d.revision):02d} | "
                            f"M:{m_val} D:{d_val} | {d.description}"
                        )
                    lines.append("")

        txt_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")

        with csv_path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=[
                    "mmm",
                    "machine_name",
                    "gggg",
                    "group_name",
                    "code",
                    "doc_type",
                    "state",
                    "revision",
                    "vvv",
                    "seq",
                    "description",
                    "model_exists_m",
                    "drawing_exists_d",
                    "best_model_path",
                    "best_drawing_path",
                ],
            )
            writer.writeheader()
            for mmm in sorted(machine_names.keys()):
                m_name = (machine_names.get(mmm) or "").strip()
                group_map = groups_by_machine.get(mmm, {})
                for gggg in sorted(group_map.keys()):
                    g_name = (group_map.get(gggg) or "").strip()
                    dlist = docs_by_pair.get((mmm, gggg), [])
                    if not dlist:
                        writer.writerow(
                            {
                                "mmm": mmm,
                                "machine_name": m_name,
                                "gggg": gggg,
                                "group_name": g_name,
                                "code": "",
                                "doc_type": "",
                                "state": "",
                                "revision": "",
                                "vvv": "",
                                "seq": "",
                                "description": "",
                                "model_exists_m": "",
                                "drawing_exists_d": "",
                                "best_model_path": "",
                                "best_drawing_path": "",
                            }
                        )
                        continue
                    for d in dlist:
                        model_best, drw_best = self._best_model_and_drw_paths(d)
                        m_ok, d_ok = self._model_and_drawing_flags(d)
                        writer.writerow(
                            {
                                "mmm": mmm,
                                "machine_name": m_name,
                                "gggg": gggg,
                                "group_name": g_name,
                                "code": d.code,
                                "doc_type": d.doc_type,
                                "state": d.state,
                                "revision": int(d.revision),
                                "vvv": d.vvv,
                                "seq": int(d.seq),
                                "description": d.description,
                                "model_exists_m": "YES" if m_ok else "NO",
                                "drawing_exists_d": "YES" if d_ok else "NO",
                                "best_model_path": model_best,
                                "best_drawing_path": drw_best,
                            }
                        )

        info(
            "Report generale gerarchico generato.\n\n"
            f"TXT: {txt_path}\n"
            f"CSV: {csv_path}"
        )
