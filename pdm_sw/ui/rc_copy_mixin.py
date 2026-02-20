from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import messagebox

import customtkinter as ctk

from pdm_sw.models import Document
from pdm_sw.codegen import build_code, build_machine_code, build_group_code
from pdm_sw.archive import (
    archive_dirs,
    archive_dirs_for_machine,
    archive_dirs_for_group,
    model_path,
    drw_path,
    inrev_tag,
    safe_copy,
    set_readonly,
)


def warn(msg: str) -> None:
    messagebox.showwarning("PDM", msg)


def info(msg: str) -> None:
    messagebox.showinfo("PDM", msg)


class RCCopyMixin:
    def _prompt_copy_target_code(self, src_doc: Document) -> dict | None:
        result: dict[str, dict | None] = {"payload": None}
        src_type = str(getattr(src_doc, "doc_type", "") or "").strip().upper()
        is_machine = src_type == "MACHINE"
        is_group = src_type == "GROUP"
        supports_vvv = src_type in ("PART", "ASSY")

        top = ctk.CTkToplevel(self)
        top.title("Copia documento")
        top.geometry("820x380")
        top.grab_set()

        ctk.CTkLabel(top, text="Nuovo codice per copia", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", padx=12, pady=(12, 6))
        ctk.CTkLabel(
            top,
            text=f"Sorgente: {src_doc.code} | Stato: {src_doc.state} | Tipo: {src_doc.doc_type}",
            text_color="#555555",
        ).pack(anchor="w", padx=12, pady=(0, 10))

        frm = ctk.CTkFrame(top)
        frm.pack(fill="x", padx=12, pady=(0, 10))

        mmm_var = tk.StringVar(value=str(src_doc.mmm or ""))
        gggg_var = tk.StringVar(value=str(src_doc.gggg or ""))
        src_vvv = str(src_doc.vvv or "").strip()
        default_use_vvv = bool(src_vvv) if src_vvv else bool(getattr(self.cfg.code, "include_vvv_by_default", False))
        use_vvv_var = tk.BooleanVar(value=(default_use_vvv if supports_vvv else False))
        vvv_var = tk.StringVar(value=src_vvv)
        preview_var = tk.StringVar(value="Prossimo codice: -")
        gggg_menu = None
        vvv_menu = None

        def _normalize_for_segment(seg: str, value: str) -> str:
            try:
                return self._norm_segment(seg, value)
            except Exception:
                return str(value or "").strip()

        def _unique(values: list[str], seg: str = "") -> list[str]:
            out: list[str] = []
            seen: set[str] = set()
            for v in values:
                raw = str(v or "").strip()
                vv = _normalize_for_segment(seg, raw) if seg else raw
                key = vv.casefold()
                if not vv or key in seen:
                    continue
                seen.add(key)
                out.append(vv)
            return out

        def _mmm_values() -> list[str]:
            vals: list[str] = []
            try:
                vals.extend([str(m[0]).strip() for m in self.store.list_machines() if str(m[0]).strip()])
            except Exception:
                pass
            vals.extend([str(src_doc.mmm or "").strip()])
            vals = _unique(vals, "MMM")
            return vals or [_normalize_for_segment("MMM", str(src_doc.mmm or "").strip()) or ""]

        def _gggg_values(mmm_sel: str) -> list[str]:
            if is_machine:
                return [""]
            vals: list[str] = []
            m_sel = _normalize_for_segment("MMM", mmm_sel)
            try:
                vals.extend([str(g[0]).strip() for g in self.store.list_groups(m_sel) if str(g[0]).strip()])
            except Exception:
                pass
            if m_sel == _normalize_for_segment("MMM", str(src_doc.mmm or "").strip()):
                vals.append(str(src_doc.gggg or "").strip())
            vals = _unique(vals, "GGGG")
            return vals or [_normalize_for_segment("GGGG", str(src_doc.gggg or "").strip()) or ""]

        def _vvv_values(mmm_sel: str, gggg_sel: str) -> list[str]:
            if not supports_vvv:
                return [""]
            vals: list[str] = []
            vals.extend([str(v).strip() for v in (self.cfg.code.vvv_presets or []) if str(v).strip()])
            vals.append(src_vvv)

            m_sel = _normalize_for_segment("MMM", mmm_sel)
            g_sel = _normalize_for_segment("GGGG", gggg_sel)
            try:
                docs = self.store.search_documents(
                    mmm=m_sel if m_sel else "",
                    gggg=g_sel if g_sel else "",
                    include_obs=True,
                )
            except Exception:
                docs = []
            for d in docs:
                vv = str(getattr(d, "vvv", "") or "").strip()
                if vv:
                    vals.append(vv)

            vals = _unique(vals, "VVV")
            if vals:
                return vals
            fallback = src_vvv or (self.cfg.code.vvv_presets[0] if self.cfg.code.vvv_presets else "V01")
            return [_normalize_for_segment("VVV", fallback)]

        mmm_values = _mmm_values()
        if _normalize_for_segment("MMM", mmm_var.get()) not in mmm_values:
            mmm_var.set(mmm_values[0])
        else:
            mmm_var.set(_normalize_for_segment("MMM", mmm_var.get()))

        gggg_values = _gggg_values(mmm_var.get())
        if _normalize_for_segment("GGGG", gggg_var.get()) not in gggg_values:
            gggg_var.set(gggg_values[0])
        else:
            gggg_var.set(_normalize_for_segment("GGGG", gggg_var.get()))

        vvv_values = _vvv_values(mmm_var.get(), gggg_var.get())
        if _normalize_for_segment("VVV", vvv_var.get()) not in vvv_values:
            vvv_var.set(vvv_values[0])
        else:
            vvv_var.set(_normalize_for_segment("VVV", vvv_var.get()))

        row1 = ctk.CTkFrame(frm, fg_color="transparent")
        row1.pack(fill="x", padx=8, pady=(8, 4))
        ctk.CTkLabel(row1, text="MMM", width=50, anchor="w").pack(side="left")
        mmm_menu = ctk.CTkOptionMenu(row1, variable=mmm_var, values=mmm_values, width=110)
        mmm_menu.pack(side="left", padx=(0, 12))
        if not is_machine:
            ctk.CTkLabel(row1, text="GGGG", width=50, anchor="w").pack(side="left")
            gggg_menu = ctk.CTkOptionMenu(row1, variable=gggg_var, values=gggg_values, width=130)
            gggg_menu.pack(side="left", padx=(0, 12))
        if supports_vvv:
            ctk.CTkCheckBox(row1, text="Usa VVV", variable=use_vvv_var, command=lambda: _on_vvv_toggle()).pack(side="left", padx=(6, 8))
            ctk.CTkLabel(row1, text="VVV", width=40, anchor="w").pack(side="left")
            vvv_menu = ctk.CTkOptionMenu(row1, variable=vvv_var, values=vvv_values, width=110)
            vvv_menu.pack(side="left", padx=(0, 4))

        row2 = ctk.CTkFrame(frm, fg_color="transparent")
        row2.pack(fill="x", padx=8, pady=(2, 8))
        ctk.CTkLabel(row2, textvariable=preview_var, font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        ctk.CTkButton(row2, text="AGGIORNA PREVIEW", width=180, command=lambda: _refresh_preview()).pack(side="left", padx=10)

        def _refresh_gggg_choices():
            if gggg_menu is None:
                return
            vals = _gggg_values(mmm_var.get())
            gggg_menu.configure(values=vals)
            cur = _normalize_for_segment("GGGG", gggg_var.get())
            if cur not in vals:
                gggg_var.set(vals[0])
            else:
                gggg_var.set(cur)

        def _refresh_vvv_choices():
            if vvv_menu is None:
                return
            vals = _vvv_values(mmm_var.get(), gggg_var.get())
            vvv_menu.configure(values=vals)
            cur = _normalize_for_segment("VVV", vvv_var.get())
            if cur not in vals:
                vvv_var.set(vals[0])
            else:
                vvv_var.set(cur)

        def _on_mmm_changed(_choice: str | None = None):
            if not is_machine:
                _refresh_gggg_choices()
            if supports_vvv:
                _refresh_vvv_choices()
            _refresh_preview()

        def _on_gggg_changed(_choice: str | None = None):
            if supports_vvv:
                _refresh_vvv_choices()
            _refresh_preview()

        def _on_vvv_toggle():
            if vvv_menu is None:
                _refresh_preview()
                return
            try:
                vvv_menu.configure(state=("normal" if bool(use_vvv_var.get()) else "disabled"))
            except Exception:
                pass
            _refresh_preview()

        def _refresh_preview():
            mmm_n = self._norm_segment("MMM", mmm_var.get())
            if not mmm_n:
                preview_var.set("Prossimo codice: - (inserisci MMM)")
                return
            try:
                if is_machine:
                    row = self.store.conn.execute(
                        "SELECT next_ver FROM ver_counters WHERE mmm=? AND gggg='' AND doc_type='MACHINE';",
                        (mmm_n,),
                    ).fetchone()
                    seq = int(row["next_ver"]) if row else 1
                    code = build_machine_code(self.cfg, mmm_n, seq)
                elif is_group:
                    gggg_n = self._norm_segment("GGGG", gggg_var.get())
                    if not gggg_n:
                        preview_var.set("Prossimo codice: - (inserisci MMM e GGGG)")
                        return
                    row = self.store.conn.execute(
                        "SELECT next_ver FROM ver_counters WHERE mmm=? AND gggg=? AND doc_type='GROUP';",
                        (mmm_n, gggg_n),
                    ).fetchone()
                    seq = int(row["next_ver"]) if row else 1
                    code = build_group_code(self.cfg, mmm_n, gggg_n, seq)
                else:
                    gggg_n = self._norm_segment("GGGG", gggg_var.get())
                    use_vvv = bool(use_vvv_var.get())
                    vvv_n = self._norm_segment("VVV", vvv_var.get()) if use_vvv else ""
                    if not gggg_n:
                        preview_var.set("Prossimo codice: - (inserisci MMM e GGGG)")
                        return
                    seq = self.store.peek_seq(mmm_n, gggg_n, vvv_n, src_doc.doc_type)
                    code = build_code(self.cfg, mmm_n, gggg_n, seq, vvv=vvv_n, force_vvv=use_vvv)
                preview_var.set(f"Prossimo codice: {code}")
            except Exception as e:
                preview_var.set(f"Preview non disponibile: {e}")

        def _cancel():
            result["payload"] = None
            top.destroy()

        def _ok():
            mmm = self._validate_segment_strict("MMM", mmm_var.get(), "MMM")
            if mmm is None:
                return
            vvv = ""
            if is_machine:
                gggg = ""
                use_vvv = False
            elif is_group:
                gggg = self._validate_segment_strict("GGGG", gggg_var.get(), "GGGG")
                if gggg is None:
                    return
                use_vvv = False
            else:
                gggg = self._validate_segment_strict("GGGG", gggg_var.get(), "GGGG")
                if gggg is None:
                    return
                use_vvv = bool(use_vvv_var.get())
                if use_vvv:
                    vvv_ok = self._validate_segment_strict("VVV", vvv_var.get(), "VVV")
                    if vvv_ok is None:
                        return
                    vvv = vvv_ok

            result["payload"] = {
                "mmm": mmm,
                "gggg": gggg,
                "vvv": vvv,
                "use_vvv": use_vvv,
                "doc_type": src_type,
            }
            top.destroy()

        mmm_menu.configure(command=_on_mmm_changed)
        if gggg_menu is not None:
            gggg_menu.configure(command=_on_gggg_changed)
        if vvv_menu is not None:
            vvv_menu.configure(command=lambda _c=None: _refresh_preview())
            _on_vvv_toggle()
        else:
            _refresh_preview()
        _refresh_preview()

        btns = ctk.CTkFrame(top, fg_color="transparent")
        btns.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(btns, text="Annulla", width=120, command=_cancel).pack(side="right", padx=6)
        ctk.CTkButton(btns, text="Conferma", width=120, command=_ok).pack(side="right", padx=6)

        top.bind("<Escape>", lambda _e: _cancel())
        try:
            mmm_menu.focus_set()
        except Exception:
            pass
        self.wait_window(top)
        return result["payload"]

    def _resolve_copy_source_paths(self, src_doc: Document) -> tuple[Path, Path | None]:
        state = str(getattr(src_doc, "state", "") or "").strip().upper()
        if state not in ("WIP", "REL"):
            raise ValueError("La copia e consentita solo da WIP o REL.")
        src_type = str(getattr(src_doc, "doc_type", "") or "").strip().upper()

        root = str(getattr(self.cfg.solidworks, "archive_root", "") or "").strip()
        wip = rel = None
        if root:
            try:
                if src_type == "MACHINE":
                    wip, rel, _inrev, _rev = archive_dirs_for_machine(root, src_doc.mmm)
                elif src_type == "GROUP":
                    wip, rel, _inrev, _rev = archive_dirs_for_group(root, src_doc.mmm, src_doc.gggg)
                else:
                    wip, rel, _inrev, _rev = archive_dirs(root, src_doc.mmm, src_doc.gggg)
            except Exception:
                wip = rel = None

        if state == "WIP":
            model_s = str(src_doc.file_wip_path or "")
            drw_s = str(src_doc.file_wip_drw_path or "")
            if not model_s and wip is not None:
                model_s = str(model_path(wip, src_doc.code, src_doc.doc_type))
            if not drw_s and wip is not None:
                drw_s = str(drw_path(wip, src_doc.code))
        else:
            model_s = str(src_doc.file_rel_path or "")
            drw_s = str(src_doc.file_rel_drw_path or "")
            if not model_s and rel is not None:
                model_s = str(model_path(rel, src_doc.code, src_doc.doc_type))
            if not drw_s and rel is not None:
                drw_s = str(drw_path(rel, src_doc.code))

        if not model_s:
            raise FileNotFoundError("Percorso modello sorgente non disponibile.")
        src_model = Path(model_s)
        if not src_model.exists():
            raise FileNotFoundError(f"Modello sorgente non trovato: {src_model}")

        src_drw = None
        candidates: list[Path] = []
        if drw_s:
            candidates.append(Path(drw_s))
        candidates.append(src_model.with_suffix(".slddrw"))
        for cand in candidates:
            try:
                if cand.exists():
                    src_drw = cand
                    break
            except Exception:
                continue

        return src_model, src_drw

    def _copy_selected_code_to_new_wip(self):
        code = self._get_table_selected_code()
        if not code:
            warn("Seleziona un codice in Ricerca&Consultazione.")
            return

        src_doc = self.store.get_document(code)
        if not src_doc:
            warn("Documento sorgente non trovato.")
            return

        src_state = str(getattr(src_doc, "state", "") or "").strip().upper()
        if src_state not in ("WIP", "REL"):
            warn("La copia e consentita solo da documenti in stato WIP o REL.")
            return

        archive_root = str(getattr(self.cfg.solidworks, "archive_root", "") or "").strip()
        if not archive_root:
            warn("Archivio non impostato (tab SolidWorks).")
            return

        target = self._prompt_copy_target_code(src_doc)
        if not target:
            return

        src_type = str(getattr(src_doc, "doc_type", "") or "").strip().upper()
        new_mmm = str(target.get("mmm") or "").strip()
        new_gggg = str(target.get("gggg") or "").strip()
        new_vvv = str(target.get("vvv") or "").strip()
        new_use_vvv = bool(target.get("use_vvv"))
        if src_type == "MACHINE":
            new_gggg = ""
            new_vvv = ""
            new_use_vvv = False
        elif src_type == "GROUP":
            new_vvv = ""
            new_use_vvv = False

        if not new_mmm or (src_type != "MACHINE" and not new_gggg):
            warn("Dati nuovo codice non validi.")
            return

        try:
            src_model, src_drw = self._resolve_copy_source_paths(src_doc)
        except Exception as e:
            warn(f"Impossibile leggere il sorgente: {e}")
            return

        try:
            if src_type == "MACHINE":
                new_seq = self.store.allocate_ver_seq(new_mmm, "", "MACHINE")
                new_code = build_machine_code(self.cfg, new_mmm, new_seq)
            elif src_type == "GROUP":
                new_seq = self.store.allocate_ver_seq(new_mmm, new_gggg, "GROUP")
                new_code = build_group_code(self.cfg, new_mmm, new_gggg, new_seq)
            else:
                new_seq = self.store.allocate_seq(new_mmm, new_gggg, new_vvv, src_doc.doc_type)
                new_code = build_code(self.cfg, new_mmm, new_gggg, new_seq, vvv=new_vvv, force_vvv=new_use_vvv)
        except Exception as e:
            warn(f"Impossibile allocare nuovo codice: {e}")
            return

        if self.store.get_document(new_code):
            warn(f"Il codice {new_code} esiste gia nel database.")
            return
        if new_code == src_doc.code:
            warn("Il nuovo codice coincide con il sorgente.")
            return

        if src_type == "MACHINE":
            wip, rel, inrev, rev = archive_dirs_for_machine(archive_root, new_mmm)
        elif src_type == "GROUP":
            wip, rel, inrev, rev = archive_dirs_for_group(archive_root, new_mmm, new_gggg)
        else:
            wip, rel, inrev, rev = archive_dirs(archive_root, new_mmm, new_gggg)
        _ = rev
        _ = rel
        _ = inrev
        dst_model = model_path(wip, new_code, src_doc.doc_type)
        dst_drw = drw_path(wip, new_code)

        if dst_model.exists():
            warn(f"Modello destinazione gia presente:\n{dst_model}")
            return
        if src_drw is not None and dst_drw.exists():
            warn(f"Disegno destinazione gia presente:\n{dst_drw}")
            return

        copied_paths: list[Path] = []
        drw_copied = False
        warnings: list[str] = []

        try:
            safe_copy(src_model, dst_model, overwrite=False)
            copied_paths.append(dst_model)
            set_readonly(dst_model, readonly=False)

            if src_drw is not None:
                safe_copy(src_drw, dst_drw, overwrite=False)
                copied_paths.append(dst_drw)
                set_readonly(dst_drw, readonly=False)
                drw_copied = True
            else:
                warnings.append("Disegno sorgente non trovato: copiato solo il modello.")

            new_doc = Document(
                id=0,
                code=new_code,
                doc_type=src_doc.doc_type,
                mmm=new_mmm,
                gggg=new_gggg,
                seq=new_seq,
                vvv=(new_vvv if src_type in ("PART", "ASSY") else ""),
                revision=0,
                state="WIP",
                obs_prev_state="",
                description=str(src_doc.description or ""),
                file_wip_path=str(dst_model),
                file_rel_path="",
                file_inrev_path="",
                file_wip_drw_path=str(dst_drw) if drw_copied else "",
                file_rel_drw_path="",
                file_inrev_drw_path="",
                created_at="",
                updated_at="",
            )
            self.store.add_document(new_doc)

            try:
                custom_vals = self.store.get_custom_values(src_doc.code) or {}
            except Exception:
                custom_vals = {}
            for prop_name, prop_value in custom_vals.items():
                self.store.set_custom_value(new_code, str(prop_name), str(prop_value))

            self._workflow_log_line(
                "COPY | "
                f"src={src_doc.code} ({src_state}) -> dst={new_code} (WIP) | "
                f"type={src_doc.doc_type} | mmm={new_mmm} gggg={new_gggg} vvv={new_vvv or '-'} | "
                f"model={dst_model} | drw={'YES' if drw_copied else 'NO'} | custom_props={len(custom_vals)}"
            )

        except Exception as e:
            for p in reversed(copied_paths):
                try:
                    if p.exists():
                        p.unlink()
                except Exception:
                    pass
            warn(f"Copia fallita: {e}")
            return

        self._refresh_all()
        info(
            "Copia completata.\n\n"
            f"Sorgente: {src_doc.code} ({src_state})\n"
            f"Nuovo codice: {new_code}\n"
            "Nuovo stato: WIP | Rev: 00\n"
            f"Proprieta custom copiate: {len(custom_vals)}"
        )
        if warnings:
            warn("\n".join(warnings))
