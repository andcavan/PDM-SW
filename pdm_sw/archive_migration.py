from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional

from .archive import (
    IN_REV_DIR,
    REV_DIR,
    drw_ext,
    drw_path,
    ext_for_doc_type,
    inrev_tag,
    model_path,
    safe_move,
)
from .models import Document
from .store import Store


@dataclass
class _MoveItem:
    src: Path
    dst: Path
    code: str
    reason: str


def _norm_path(p: Path | str) -> str:
    try:
        return str(Path(p).resolve()).lower()
    except Exception:
        return str(p).replace("/", "\\").lower()


def _existing_unique(paths: List[Path]) -> List[Path]:
    out: List[Path] = []
    seen: set[str] = set()
    for p in paths:
        try:
            if not p.exists():
                continue
        except Exception:
            continue
        k = _norm_path(p)
        if k in seen:
            continue
        seen.add(k)
        out.append(p)
    return out


def _doc_dirs(archive_root: str, doc: Document) -> tuple[Path, Path, Path]:
    root = Path(archive_root)
    dt = str(doc.doc_type or "").upper()
    if dt == "MACHINE":
        current = root / doc.mmm
    else:
        current = root / doc.mmm / doc.gggg
    inrev = current / IN_REV_DIR
    rev = current / REV_DIR
    return current, inrev, rev


def _old_bases(archive_root: str, doc: Document) -> List[Path]:
    root = Path(archive_root)
    dt = str(doc.doc_type or "").upper()
    if dt == "MACHINE":
        return [root / "MACHINES" / doc.mmm, root / doc.mmm]
    if dt == "GROUP":
        return [root / "GROUPS" / doc.mmm / doc.gggg, root / doc.mmm / doc.gggg]
    return [root / doc.mmm / doc.gggg]


def _pick_current_source(doc: Document, old_bases: List[Path], is_drw: bool = False) -> Optional[Path]:
    if is_drw:
        explicit = [Path(p) for p in [doc.file_wip_drw_path, doc.file_rel_drw_path] if str(p or "").strip()]
        ext = drw_ext()
    else:
        explicit = [Path(p) for p in [doc.file_wip_path, doc.file_rel_path] if str(p or "").strip()]
        ext = ext_for_doc_type(doc.doc_type)

    name = f"{doc.code}{ext}"
    fallback: List[Path] = []
    for b in old_bases:
        fallback.extend(
            [
                b / name,
                b / "wip" / name,
                b / "WIP" / name,
                b / "rel" / name,
                b / "REL" / name,
            ]
        )
    existing = _existing_unique(explicit + fallback)
    return existing[0] if existing else None


def _pick_inrev_source(doc: Document, old_bases: List[Path], is_drw: bool = False) -> Optional[Path]:
    explicit = Path(doc.file_inrev_drw_path if is_drw else doc.file_inrev_path) if (
        doc.file_inrev_drw_path if is_drw else doc.file_inrev_path
    ) else None
    if explicit and explicit.exists():
        return explicit

    suffix = drw_ext() if is_drw else ext_for_doc_type(doc.doc_type)
    found: List[Path] = []
    for b in old_bases:
        for folder in ("inrev", "IN_REV"):
            d = b / folder
            if not d.exists():
                continue
            patt = f"{doc.code}_R*__INREV{suffix}"
            for p in sorted(d.glob(patt), key=lambda x: x.name, reverse=True):
                if p.is_file():
                    found.append(p)
    return found[0] if found else None


def _default_inrev_name(doc: Document, is_drw: bool = False) -> str:
    tag = inrev_tag(doc.code, int(doc.revision))
    return f"{tag}{drw_ext() if is_drw else ext_for_doc_type(doc.doc_type)}"


def _target_from_source_or_default(inrev_dir: Path, source: Optional[Path], default_name: str) -> Path:
    if source is not None:
        return inrev_dir / source.name
    return inrev_dir / default_name


def _collect_history_moves(
    doc: Document,
    old_bases: List[Path],
    inrev_dir: Path,
    rev_dir: Path,
) -> List[_MoveItem]:
    out: List[_MoveItem] = []
    model_ext = ext_for_doc_type(doc.doc_type)
    drawing_ext = drw_ext()

    for b in old_bases:
        for folder in ("inrev", "IN_REV"):
            d = b / folder
            if not d.exists():
                continue
            for patt in (f"{doc.code}_R*__INREV{model_ext}", f"{doc.code}_R*__INREV{drawing_ext}"):
                for src in d.glob(patt):
                    if src.is_file():
                        out.append(_MoveItem(src=src, dst=inrev_dir / src.name, code=doc.code, reason="INREV_HISTORY"))

        for folder in ("rev", "REV"):
            d = b / folder
            if not d.exists():
                continue
            for patt in (f"{doc.code}_R*{model_ext}", f"{doc.code}_R*{drawing_ext}"):
                for src in d.glob(patt):
                    if src.is_file():
                        out.append(_MoveItem(src=src, dst=rev_dir / src.name, code=doc.code, reason="REV_HISTORY"))

    return out


def run_archive_layout_migration(
    store: Store,
    archive_root: str,
    apply_changes: bool = False,
) -> Dict[str, Any]:
    root = Path(str(archive_root or "").strip())
    if not str(root).strip():
        return {"ok": False, "error": "Archivio non configurato.", "apply_changes": bool(apply_changes)}

    docs = store.list_documents(include_obs=True)
    move_map: Dict[str, _MoveItem] = {}
    conflicts: List[str] = []
    updates_by_code: Dict[str, Dict[str, str]] = {}

    def add_move(item: _MoveItem) -> None:
        if _norm_path(item.src) == _norm_path(item.dst):
            return
        k = _norm_path(item.src)
        prev = move_map.get(k)
        if prev is None:
            move_map[k] = item
            return
        if _norm_path(prev.dst) != _norm_path(item.dst):
            conflicts.append(
                f"{item.code}: sorgente duplicata con target diversi | {item.src} -> {prev.dst} / {item.dst}"
            )

    for doc in docs:
        current_dir, inrev_dir, rev_dir = _doc_dirs(str(root), doc)
        old_bases = _old_bases(str(root), doc)

        current_model_target = model_path(current_dir, doc.code, doc.doc_type)
        current_drw_target = drw_path(current_dir, doc.code)

        src_current_model = _pick_current_source(doc, old_bases, is_drw=False)
        src_current_drw = _pick_current_source(doc, old_bases, is_drw=True)
        src_inrev_model = _pick_inrev_source(doc, old_bases, is_drw=False)
        src_inrev_drw = _pick_inrev_source(doc, old_bases, is_drw=True)

        inrev_model_target = _target_from_source_or_default(
            inrev_dir,
            src_inrev_model,
            _default_inrev_name(doc, is_drw=False),
        )
        inrev_drw_target = _target_from_source_or_default(
            inrev_dir,
            src_inrev_drw,
            _default_inrev_name(doc, is_drw=True),
        )

        if src_current_model is not None:
            add_move(_MoveItem(src=src_current_model, dst=current_model_target, code=doc.code, reason="CURRENT_MODEL"))
        if src_current_drw is not None:
            add_move(_MoveItem(src=src_current_drw, dst=current_drw_target, code=doc.code, reason="CURRENT_DRW"))
        if src_inrev_model is not None:
            add_move(_MoveItem(src=src_inrev_model, dst=inrev_model_target, code=doc.code, reason="INREV_MODEL"))
        if src_inrev_drw is not None:
            add_move(_MoveItem(src=src_inrev_drw, dst=inrev_drw_target, code=doc.code, reason="INREV_DRW"))

        for h in _collect_history_moves(doc, old_bases, inrev_dir=inrev_dir, rev_dir=rev_dir):
            add_move(h)

        updates: Dict[str, str] = {}
        has_current_model_ref = bool(str(doc.file_wip_path or "").strip() or str(doc.file_rel_path or "").strip() or src_current_model)
        has_current_drw_ref = bool(str(doc.file_wip_drw_path or "").strip() or str(doc.file_rel_drw_path or "").strip() or src_current_drw)
        has_inrev_model_ref = bool(str(doc.file_inrev_path or "").strip() or src_inrev_model)
        has_inrev_drw_ref = bool(str(doc.file_inrev_drw_path or "").strip() or src_inrev_drw)

        if has_current_model_ref:
            updates["file_wip_path"] = str(current_model_target)
            updates["file_rel_path"] = str(current_model_target)
        if has_current_drw_ref:
            updates["file_wip_drw_path"] = str(current_drw_target)
            updates["file_rel_drw_path"] = str(current_drw_target)
        if has_inrev_model_ref:
            updates["file_inrev_path"] = str(inrev_model_target)
        if has_inrev_drw_ref:
            updates["file_inrev_drw_path"] = str(inrev_drw_target)

        if updates:
            updates_by_code[doc.code] = updates

    moves = list(move_map.values())
    moved = 0
    skipped_missing = 0
    move_errors: List[str] = []
    runtime_conflicts = list(conflicts)

    if apply_changes:
        for m in moves:
            try:
                if not m.src.exists():
                    skipped_missing += 1
                    continue
                if m.dst.exists() and _norm_path(m.src) != _norm_path(m.dst):
                    runtime_conflicts.append(f"{m.code}: target gia presente | {m.dst}")
                    continue
                safe_move(m.src, m.dst, overwrite=False)
                moved += 1
            except FileExistsError:
                runtime_conflicts.append(f"{m.code}: target gia presente | {m.dst}")
            except Exception as e:
                move_errors.append(f"{m.code}: move error {m.src} -> {m.dst} | {e}")

        docs_updated = 0
        for code, fields in updates_by_code.items():
            try:
                store.update_document(code, **fields)
                docs_updated += 1
            except Exception as e:
                move_errors.append(f"{code}: update_document error | {e}")
    else:
        docs_updated = len(updates_by_code)

    ok = (len(move_errors) == 0)
    return {
        "ok": ok,
        "apply_changes": bool(apply_changes),
        "archive_root": str(root),
        "docs_scanned": len(docs),
        "docs_to_update": len(updates_by_code),
        "docs_updated": int(docs_updated),
        "moves_planned": len(moves),
        "moves_done": int(moved),
        "moves_missing": int(skipped_missing),
        "conflicts": runtime_conflicts,
        "errors": move_errors,
        "sample_moves": [f"{m.code} | {m.reason} | {m.src} -> {m.dst}" for m in moves[:25]],
    }
