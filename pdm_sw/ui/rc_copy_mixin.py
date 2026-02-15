from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import messagebox

import customtkinter as ctk

from pdm_sw.models import Document
from pdm_sw.codegen import build_code
from pdm_sw.archive import archive_dirs, model_path, drw_path, inrev_tag, safe_copy, set_readonly


def warn(msg: str) -> None:
    messagebox.showwarning("PDM", msg)


def info(msg: str) -> None:
    messagebox.showinfo("PDM", msg)


class RCCopyMixin:
    def _prompt_copy_target_code(self, src_doc: Document) -> dict | None:
        result: dict[str, dict | None] = {"payload": None}

        top = ctk.CTkToplevel(self)
        top.title("Copia documento")
        top.geometry("760x360")
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
        use_vvv_var = tk.BooleanVar(value=bool(src_vvv))
        vvv_var = tk.StringVar(value=src_vvv)
        preview_var = tk.StringVar(value="Prossimo codice: -")

        def _unique(values: list[str]) -> list[str]:
            out: list[str] = []
            seen: set[str] = set()
            for v in values:
                vv = str(v or "").strip().upper()
                if not vv or vv in seen:
                    continue
                seen.add(vv)
                out.append(vv)
            return out

        def _mmm_values() -> list[str]:
            vals: list[str] = []
            try:
                vals.extend([str(m[0]).strip().upper() for m in self.store.list_machines() if str(m[0]).strip()])
            except Exception:
                pass
            vals.extend([str(src_doc.mmm or "").strip().upper()])
            vals = _unique(vals)
            return vals or [str(src_doc.mmm or "").strip().upper() or ""]

        def _gggg_values(mmm_sel: str) -> list[str]:
            vals: list[str] = []
            m_sel = str(mmm_sel or "").strip().upper()
            try:
                vals.extend([str(g[0]).strip().upper() for g in self.store.list_groups(m_sel) if str(g[0]).strip()])
            except Exception:
                pass
            if m_sel == str(src_doc.mmm or "").strip().upper():
                vals.append(str(src_doc.gggg or "").strip().upper())
            vals = _unique(vals)
            return vals or [str(src_doc.gggg or "").strip().upper() or ""]

        def _vvv_values(mmm_sel: str, gggg_sel: str) -> list[str]:
            vals: list[str] = []
            vals.extend([str(v).strip().upper() for v in (self.cfg.code.vvv_presets or []) if str(v).strip()])
            vals.append(src_vvv.upper())

            m_sel = str(mmm_sel or "").strip().upper()
            g_sel = str(gggg_sel or "").strip().upper()
            try:
                docs = self.store.search_documents(
                    mmm=m_sel if m_sel else "",
                    gggg=g_sel if g_sel else "",
                    include_obs=True,
                )
            except Exception:
                docs = []
            for d in docs:
                vv = str(getattr(d, "vvv", "") or "").strip().upper()
                if vv:
                    vals.append(vv)

            vals = _unique(vals)
            if vals:
                return vals
            return [src_vvv.upper() or "V01"]

        mmm_values = _mmm_values()
        if str(mmm_var.get() or "").strip().upper() not in mmm_values:
            mmm_var.set(mmm_values[0])
        else:
            mmm_var.set(str(mmm_var.get() or "").strip().upper())

        gggg_values = _gggg_values(mmm_var.get())
        if str(gggg_var.get() or "").strip().upper() not in gggg_values:
            gggg_var.set(gggg_values[0])
        else:
            gggg_var.set(str(gggg_var.get() or "").strip().upper())

        vvv_values = _vvv_values(mmm_var.get(), gggg_var.get())
        if str(vvv_var.get() or "").strip().upper() not in vvv_values:
            vvv_var.set(vvv_values[0])
        else:
            vvv_var.set(str(vvv_var.get() or "").strip().upper())

        row1 = ctk.CTkFrame(frm, fg_color="transparent")
        row1.pack(fill="x", padx=8, pady=(8, 4))
        ctk.CTkLabel(row1, text="MMM", width=50, anchor="w").pack(side="left")
        mmm_menu = ctk.CTkOptionMenu(row1, variable=mmm_var, values=mmm_values, width=110)
        mmm_menu.pack(side="left", padx=(0, 12))
        ctk.CTkLabel(row1, text="GGGG", width=50, anchor="w").pack(side="left")
        gggg_menu = ctk.CTkOptionMenu(row1, variable=gggg_var, values=gggg_values, width=130)
        gggg_menu.pack(side="left", padx=(0, 12))
        ctk.CTkCheckBox(row1, text="Usa VVV", variable=use_vvv_var, command=lambda: _on_vvv_toggle()).pack(side="left", padx=(6, 8))
        ctk.CTkLabel(row1, text="VVV", width=40, anchor="w").pack(side="left")
        vvv_menu = ctk.CTkOptionMenu(row1, variable=vvv_var, values=vvv_values, width=110)
        vvv_menu.pack(side="left", padx=(0, 4))

        row2 = ctk.CTkFrame(frm, fg_color="transparent")
        row2.pack(fill="x", padx=8, pady=(2, 8))
        ctk.CTkLabel(row2, textvariable=preview_var, font=ctk.CTkFont(size=13, weight="bold")).pack(side="left")
        ctk.CTkButton(row2, text="AGGIORNA PREVIEW", width=180, command=lambda: _refresh_preview()).pack(side="left", padx=10)

        def _refresh_gggg_choices():
            vals = _gggg_values(mmm_var.get())
            gggg_menu.configure(values=vals)
            cur = str(gggg_var.get() or "").strip().upper()
            if cur not in vals:
                gggg_var.set(vals[0])
            else:
                gggg_var.set(cur)

        def _refresh_vvv_choices():
            vals = _vvv_values(mmm_var.get(), gggg_var.get())
            vvv_menu.configure(values=vals)
            cur = str(vvv_var.get() or "").strip().upper()
            if cur not in vals:
                vvv_var.set(vals[0])
            else:
                vvv_var.set(cur)

        def _on_mmm_changed(_choice: str | None = None):
            _refresh_gggg_choices()
            _refresh_vvv_choices()
            _refresh_preview()

        def _on_gggg_changed(_choice: str | None = None):
            _refresh_vvv_choices()
            _refresh_preview()

        def _on_vvv_toggle():
            try:
                vvv_menu.configure(state=("normal" if bool(use_vvv_var.get()) else "disabled"))
            except Exception:
                pass
            _refresh_preview()

        def _refresh_preview():
            mmm_n = self._norm_segment("MMM", mmm_var.get())
            gggg_n = self._norm_segment("GGGG", gggg_var.get())
            use_vvv = bool(use_vvv_var.get())
            vvv_n = self._norm_segment("VVV", vvv_var.get()) if use_vvv else ""

            if not mmm_n or not gggg_n:
                preview_var.set("Prossimo codice: - (inserisci MMM e GGGG)")
                return
            try:
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
            gggg = self._validate_segment_strict("GGGG", gggg_var.get(), "GGGG")
            if gggg is None:
                return

            use_vvv = bool(use_vvv_var.get())
            vvv = ""
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
            }
            top.destroy()

        mmm_menu.configure(command=_on_mmm_changed)
        gggg_menu.configure(command=_on_gggg_changed)
        vvv_menu.configure(command=lambda _c=None: _refresh_preview())
        _on_vvv_toggle()
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

        root = str(getattr(self.cfg.solidworks, "archive_root", "") or "").strip()
        wip = rel = None
        if root:
            try:
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

        new_mmm = str(target.get("mmm") or "").strip()
        new_gggg = str(target.get("gggg") or "").strip()
        new_vvv = str(target.get("vvv") or "").strip()
        new_use_vvv = bool(target.get("use_vvv"))
        if not new_mmm or not new_gggg:
            warn("Dati nuovo codice non validi.")
            return

        try:
            src_model, src_drw = self._resolve_copy_source_paths(src_doc)
        except Exception as e:
            warn(f"Impossibile leggere il sorgente: {e}")
            return

        try:
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
                vvv=new_vvv,
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
