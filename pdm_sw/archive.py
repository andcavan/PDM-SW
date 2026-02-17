from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Tuple, Optional
import os
import shutil
import stat
import time
from datetime import datetime

from .models import Document, DocType


def ext_for_doc_type(doc_type: DocType) -> str:
    if doc_type == "PART":
        return ".sldprt"
    elif doc_type in ("ASSY", "MACHINE", "GROUP"):
        return ".sldasm"
    return ".sldasm"


def drw_ext() -> str:
    return ".slddrw"


def ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def set_readonly(path: Path, readonly: bool = True) -> None:
    if not path.exists():
        return
    try:
        mode = path.stat().st_mode
        if readonly:
            path.chmod(mode & ~stat.S_IWUSR & ~stat.S_IWGRP & ~stat.S_IWOTH)
        else:
            path.chmod(mode | stat.S_IWUSR)
    except Exception:
        pass


def archive_dirs(archive_root: str, mmm: str, gggg: str) -> Tuple[Path, Path, Path, Path]:
    root = Path(archive_root)
    base = root / mmm / gggg
    wip = base / "wip"
    rel = base / "rel"
    inrev = base / "inrev"
    rev = base / "rev"
    for p in (wip, rel, inrev, rev):
        ensure_dir(p)
    return wip, rel, inrev, rev


def archive_dirs_for_machine(archive_root: str, mmm: str) -> Tuple[Path, Path, Path, Path]:
    """Cartelle archivio per MACHINE (solo MMM, senza GGGG)."""
    root = Path(archive_root)
    base = root / "MACHINES" / mmm
    wip = base / "wip"
    rel = base / "rel"
    inrev = base / "inrev"
    rev = base / "rev"
    for p in (wip, rel, inrev, rev):
        ensure_dir(p)
    return wip, rel, inrev, rev


def archive_dirs_for_group(archive_root: str, mmm: str, gggg: str) -> Tuple[Path, Path, Path, Path]:
    """Cartelle archivio per GROUP (MMM_GGGG)."""
    root = Path(archive_root)
    base = root / "GROUPS" / mmm / gggg
    wip = base / "wip"
    rel = base / "rel"
    inrev = base / "inrev"
    rev = base / "rev"
    for p in (wip, rel, inrev, rev):
        ensure_dir(p)
    return wip, rel, inrev, rev


def model_path(folder: Path, code: str, doc_type: DocType) -> Path:
    return folder / f"{code}{ext_for_doc_type(doc_type)}"


def drw_path(folder: Path, code: str) -> Path:
    return folder / f"{code}{drw_ext()}"


def inrev_tag(code: str, revision: int) -> str:
    return f"{code}_R{revision:02d}__INREV"


def rev_tag(code: str, revision: int) -> str:
    return f"{code}_R{revision:02d}"


def _append_log(log_file: str | Path | None, message: str) -> None:
    if not log_file:
        return
    try:
        p = Path(log_file)
        ensure_dir(p.parent)
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with p.open("a", encoding="utf-8") as f:
            f.write(f"{ts} | {message}\n")
    except Exception:
        pass


def _is_lock_like_oserror(e: OSError) -> bool:
    win = int(getattr(e, "winerror", 0) or 0)
    eno = int(getattr(e, "errno", 0) or 0)
    # 32/33: sharing violation, 5: access denied, 13: permission denied
    return win in (5, 32, 33) or eno in (13,)


def _run_with_retries(fn, attempts: int = 4, delay_s: float = 0.2) -> None:
    attempts = max(1, int(attempts))
    for i in range(attempts):
        try:
            fn()
            return
        except OSError as e:
            if (i < attempts - 1) and _is_lock_like_oserror(e):
                time.sleep(delay_s)
                continue
            raise


def safe_copy(src: Path, dst: Path, overwrite: bool = False, log_file: str | Path | None = None) -> None:
    if not src.exists():
        _append_log(log_file, f"FS COPY SKIP (src missing) | {src} -> {dst}")
        return
    ensure_dir(dst.parent)
    _append_log(log_file, f"FS COPY START | {src} -> {dst} | overwrite={overwrite}")
    if dst.exists():
        if not overwrite:
            _append_log(log_file, f"FS COPY FAIL (dst exists) | {dst}")
            raise FileExistsError(f"File destinazione gia esistente: {dst}")
        safe_delete(dst, strict=True, log_file=log_file)
    _run_with_retries(lambda: shutil.copy2(src, dst))
    _append_log(log_file, f"FS COPY OK | {src} -> {dst}")


def safe_move(src: Path, dst: Path, overwrite: bool = False, log_file: str | Path | None = None) -> None:
    if not src.exists():
        _append_log(log_file, f"FS MOVE SKIP (src missing) | {src} -> {dst}")
        return
    ensure_dir(dst.parent)
    try:
        if src.resolve() == dst.resolve():
            return
    except Exception:
        pass
    _append_log(log_file, f"FS MOVE START | {src} -> {dst} | overwrite={overwrite}")
    if dst.exists():
        if not overwrite:
            _append_log(log_file, f"FS MOVE FAIL (dst exists) | {dst}")
            raise FileExistsError(f"File destinazione gia esistente: {dst}")
        safe_delete(dst, strict=True, log_file=log_file)
    # Se il move richiede copy+delete (cross-volume), il sorgente deve essere scrivibile.
    set_readonly(src, False)
    try:
        _run_with_retries(lambda: os.replace(str(src), str(dst)))
        _append_log(log_file, f"FS MOVE OK (rename) | {src} -> {dst}")
        return
    except OSError as e:
        win = int(getattr(e, "winerror", 0) or 0)
        eno = int(getattr(e, "errno", 0) or 0)
        is_cross_device = (eno == getattr(os, "EXDEV", 18)) or (win == 17)
        if not is_cross_device:
            _append_log(log_file, f"FS MOVE ERROR | {src} -> {dst} | {type(e).__name__}: {e}")
            raise
    _append_log(log_file, f"FS MOVE FALLBACK copy+delete | {src} -> {dst}")
    # fallback cross-volume: copy + delete sorgente
    safe_copy(src, dst, overwrite=overwrite, log_file=log_file)
    safe_delete(src, strict=True, log_file=log_file)
    _append_log(log_file, f"FS MOVE OK (copy+delete) | {src} -> {dst}")


def safe_delete(p: Path, strict: bool = False, log_file: str | Path | None = None) -> bool:
    try:
        if not p.exists():
            _append_log(log_file, f"FS DELETE SKIP (missing) | {p}")
            return True
        _append_log(log_file, f"FS DELETE START | {p}")
        set_readonly(p, False)
        _run_with_retries(lambda: p.unlink())
        _append_log(log_file, f"FS DELETE OK | {p}")
        return True
    except Exception as e:
        _append_log(log_file, f"FS DELETE ERROR | {p} | {type(e).__name__}: {e}")
        if strict:
            raise
        return False


@dataclass
class WorkflowResult:
    ok: bool
    message: str


def release_wip(doc: Document, archive_root: str, log_file: str | Path | None = None) -> Tuple[Document, WorkflowResult]:
    _append_log(log_file, f"WF START RELEASE | code={doc.code} | state={doc.state} | rev={int(doc.revision):02d}")
    if doc.state != "WIP":
        _append_log(log_file, "WF FAIL RELEASE | stato non valido (serve WIP)")
        return doc, WorkflowResult(False, "Per il rilascio serve stato WIP.")

    if not archive_root:
        _append_log(log_file, "WF FAIL RELEASE | archivio non configurato")
        return doc, WorkflowResult(False, "Archivio non configurato (SolidWorks > Archivio).")

    wip, rel, inrev, rev = archive_dirs(archive_root, doc.mmm, doc.gggg)

    src_model = Path(doc.file_wip_path) if doc.file_wip_path else model_path(wip, doc.code, doc.doc_type)
    dst_model = model_path(rel, doc.code, doc.doc_type)

    src_drw = Path(doc.file_wip_drw_path) if doc.file_wip_drw_path else drw_path(wip, doc.code)
    dst_drw = drw_path(rel, doc.code)

    if src_model.exists():
        safe_move(src_model, dst_model, overwrite=True, log_file=log_file)
        set_readonly(dst_model, True)
        doc.file_rel_path = str(dst_model)
        doc.file_wip_path = ""
    else:
        # allow release without file
        doc.file_rel_path = str(dst_model) if doc.file_rel_path else ""

    if src_drw.exists():
        safe_move(src_drw, dst_drw, overwrite=True, log_file=log_file)
        set_readonly(dst_drw, True)
        doc.file_rel_drw_path = str(dst_drw)
        doc.file_wip_drw_path = ""
    else:
        doc.file_rel_drw_path = doc.file_rel_drw_path or ""

    doc.state = "REL"
    # first release keeps revision as is (default 0 => 00)
    _append_log(log_file, f"WF OK RELEASE | code={doc.code} | new_state={doc.state} | rev={int(doc.revision):02d}")
    return doc, WorkflowResult(True, "Rilasciato (REL).")


def create_inrev(doc: Document, archive_root: str, log_file: str | Path | None = None) -> Tuple[Document, WorkflowResult]:
    _append_log(log_file, f"WF START CREATE_INREV | code={doc.code} | state={doc.state} | rev={int(doc.revision):02d}")
    if doc.state != "REL":
        _append_log(log_file, "WF FAIL CREATE_INREV | stato non valido (serve REL)")
        return doc, WorkflowResult(False, "Per creare revisione serve stato REL.")

    if not archive_root:
        _append_log(log_file, "WF FAIL CREATE_INREV | archivio non configurato")
        return doc, WorkflowResult(False, "Archivio non configurato (SolidWorks > Archivio).")

    wip, rel, inrev, rev = archive_dirs(archive_root, doc.mmm, doc.gggg)
    tag = inrev_tag(doc.code, doc.revision)

    src_model = Path(doc.file_rel_path) if doc.file_rel_path else model_path(rel, doc.code, doc.doc_type)
    dst_model = model_path(inrev, tag, doc.doc_type)

    src_drw = Path(doc.file_rel_drw_path) if doc.file_rel_drw_path else drw_path(rel, doc.code)
    dst_drw = drw_path(inrev, tag)

    if src_model.exists():
        safe_copy(src_model, dst_model, overwrite=True, log_file=log_file)
        set_readonly(dst_model, False)
        doc.file_inrev_path = str(dst_model)
    else:
        doc.file_inrev_path = ""

    if src_drw.exists():
        safe_copy(src_drw, dst_drw, overwrite=True, log_file=log_file)
        set_readonly(dst_drw, False)
        doc.file_inrev_drw_path = str(dst_drw)
    else:
        doc.file_inrev_drw_path = ""

    doc.state = "IN_REV"
    _append_log(log_file, f"WF OK CREATE_INREV | code={doc.code} | new_state={doc.state} | rev={int(doc.revision):02d}")
    return doc, WorkflowResult(True, "Revisione creata (IN_REV).")


def approve_inrev(doc: Document, archive_root: str, log_file: str | Path | None = None) -> Tuple[Document, WorkflowResult]:
    _append_log(log_file, f"WF START APPROVE_INREV | code={doc.code} | state={doc.state} | rev={int(doc.revision):02d}")
    if doc.state != "IN_REV":
        _append_log(log_file, "WF FAIL APPROVE_INREV | stato non valido (serve IN_REV)")
        return doc, WorkflowResult(False, "Per approvare serve stato IN_REV.")

    if not archive_root:
        _append_log(log_file, "WF FAIL APPROVE_INREV | archivio non configurato")
        return doc, WorkflowResult(False, "Archivio non configurato (SolidWorks > Archivio).")

    wip, rel, inrev, rev = archive_dirs(archive_root, doc.mmm, doc.gggg)

    # move current REL to REV with rev tag
    cur_rev = doc.revision
    rel_tag = rev_tag(doc.code, cur_rev)

    rel_model = Path(doc.file_rel_path) if doc.file_rel_path else model_path(rel, doc.code, doc.doc_type)
    rel_drw = Path(doc.file_rel_drw_path) if doc.file_rel_drw_path else drw_path(rel, doc.code)
    rev_model_dst = model_path(rev, rel_tag, doc.doc_type)
    rev_drw_dst = drw_path(rev, rel_tag)

    if rel_model.exists() and rev_model_dst.exists():
        _append_log(log_file, f"WF FAIL APPROVE_INREV | revisione gia presente: {rev_model_dst}")
        return doc, WorkflowResult(False, f"Revisione gia presente in archivio REV: {rev_model_dst}")
    if rel_drw.exists() and rev_drw_dst.exists():
        _append_log(log_file, f"WF FAIL APPROVE_INREV | drawing revisione gia presente: {rev_drw_dst}")
        return doc, WorkflowResult(False, f"Disegno revisione gia presente in archivio REV: {rev_drw_dst}")

    if rel_model.exists():
        safe_move(rel_model, rev_model_dst, overwrite=False, log_file=log_file)
        set_readonly(rev_model_dst, True)
    if rel_drw.exists():
        safe_move(rel_drw, rev_drw_dst, overwrite=False, log_file=log_file)
        set_readonly(rev_drw_dst, True)

    # promote INREV copy to REL with base code name
    inrev_model = Path(doc.file_inrev_path) if doc.file_inrev_path else model_path(inrev, inrev_tag(doc.code, cur_rev), doc.doc_type)
    inrev_drw = Path(doc.file_inrev_drw_path) if doc.file_inrev_drw_path else drw_path(inrev, inrev_tag(doc.code, cur_rev))

    if inrev_model.exists():
        dst = model_path(rel, doc.code, doc.doc_type)
        safe_move(inrev_model, dst, overwrite=True, log_file=log_file)
        set_readonly(dst, True)
        doc.file_rel_path = str(dst)
    else:
        doc.file_rel_path = doc.file_rel_path or ""

    if inrev_drw.exists():
        dst = drw_path(rel, doc.code)
        safe_move(inrev_drw, dst, overwrite=True, log_file=log_file)
        set_readonly(dst, True)
        doc.file_rel_drw_path = str(dst)
    else:
        doc.file_rel_drw_path = doc.file_rel_drw_path or ""

    # increment revision
    doc.revision = cur_rev + 1
    doc.state = "REL"
    doc.file_inrev_path = ""
    doc.file_inrev_drw_path = ""
    _append_log(log_file, f"WF OK APPROVE_INREV | code={doc.code} | new_state={doc.state} | rev={int(doc.revision):02d}")
    return doc, WorkflowResult(True, f"Approvato: REL rev {doc.revision:02d}.")


def cancel_inrev(doc: Document, log_file: str | Path | None = None) -> Tuple[Document, WorkflowResult]:
    _append_log(log_file, f"WF START CANCEL_INREV | code={doc.code} | state={doc.state} | rev={int(doc.revision):02d}")
    if doc.state != "IN_REV":
        _append_log(log_file, "WF FAIL CANCEL_INREV | stato non valido (serve IN_REV)")
        return doc, WorkflowResult(False, "Per annullare serve stato IN_REV.")

    # delete inrev copies if exist (strict: non cambiare stato se non riusciamo a pulire)
    try:
        if doc.file_inrev_path:
            safe_delete(Path(doc.file_inrev_path), strict=True, log_file=log_file)
        if doc.file_inrev_drw_path:
            safe_delete(Path(doc.file_inrev_drw_path), strict=True, log_file=log_file)
    except Exception as e:
        _append_log(log_file, f"WF FAIL CANCEL_INREV | cleanup failed: {e}")
        return doc, WorkflowResult(False, f"Impossibile eliminare i file IN_REV: {e}")

    doc.file_inrev_path = ""
    doc.file_inrev_drw_path = ""
    doc.state = "REL"
    _append_log(log_file, f"WF OK CANCEL_INREV | code={doc.code} | new_state={doc.state} | rev={int(doc.revision):02d}")
    return doc, WorkflowResult(True, "Revisione annullata (REL).")



def apply_state_permissions(doc: Document, state: str) -> None:
    """Imposta permessi lettura/scrittura sui file in base allo stato."""
    # default: tutto read-only
    paths = [
        doc.file_wip_path, doc.file_rel_path, doc.file_inrev_path,
        doc.file_wip_drw_path, doc.file_rel_drw_path, doc.file_inrev_drw_path,
    ]
    for p in paths:
        if p:
            set_readonly(Path(p), True)

    if state == "WIP":
        for p in (doc.file_wip_path, doc.file_wip_drw_path):
            if p:
                set_readonly(Path(p), False)
    elif state == "IN_REV":
        for p in (doc.file_inrev_path, doc.file_inrev_drw_path):
            if p:
                set_readonly(Path(p), False)
    # REL/OBS restano read-only


def restore_obsolete(doc: Document, prev_state: str, log_file: str | Path | None = None) -> Tuple[Document, WorkflowResult]:
    _append_log(log_file, f"WF START RESTORE_OBS | code={doc.code} | state={doc.state} | prev={prev_state} | rev={int(doc.revision):02d}")
    if doc.state != "OBS":
        _append_log(log_file, "WF FAIL RESTORE_OBS | documento non in OBS")
        return doc, WorkflowResult(False, "Il documento non è in stato OBS.")
    if prev_state not in ("WIP", "REL", "IN_REV"):
        _append_log(log_file, "WF FAIL RESTORE_OBS | prev_state non valido")
        return doc, WorkflowResult(False, "Stato precedente non valido o non disponibile.")
    doc.state = prev_state
    apply_state_permissions(doc, prev_state)
    _append_log(log_file, f"WF OK RESTORE_OBS | code={doc.code} | new_state={doc.state} | rev={int(doc.revision):02d}")
    return doc, WorkflowResult(True, f"Ripristinato da OBS a {prev_state}.")

def set_obsolete(doc: Document, log_file: str | Path | None = None) -> Tuple[Document, WorkflowResult]:
    # from any state
    _append_log(log_file, f"WF START SET_OBSOLETE | code={doc.code} | state={doc.state} | rev={int(doc.revision):02d}")
    apply_state_permissions(doc, "OBS")
    doc.state = "OBS"
    _append_log(log_file, f"WF OK SET_OBSOLETE | code={doc.code} | new_state={doc.state} | rev={int(doc.revision):02d}")
    return doc, WorkflowResult(True, "Impostato OBS (obsoleto).")
