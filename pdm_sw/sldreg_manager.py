from __future__ import annotations

import os
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path


@dataclass
class RestoreOptions:
    system_options: bool = True
    toolbar_layout: bool = True
    toolbar_mode: str = "all"
    keyboard_shortcuts: bool = True
    mouse_gestures: bool = True
    menu_customizations: bool = True
    saved_views: bool = True

    def has_any_selection(self) -> bool:
        return any(
            (
                self.system_options,
                self.toolbar_layout,
                self.keyboard_shortcuts,
                self.mouse_gestures,
                self.menu_customizations,
                self.saved_views,
            )
        )


def _normalize_toolbar_mode(value: object) -> str:
    mode = str(value or "").strip().casefold()
    if mode in ("macro", "macro_only", "macro-only", "macros", "solo_macro"):
        return "macro_only"
    return "all"


def _normalize_restore_options(options: RestoreOptions | None) -> RestoreOptions:
    if options is None:
        return RestoreOptions()
    return RestoreOptions(
        system_options=bool(options.system_options),
        toolbar_layout=bool(options.toolbar_layout),
        toolbar_mode=_normalize_toolbar_mode(options.toolbar_mode),
        keyboard_shortcuts=bool(options.keyboard_shortcuts),
        mouse_gestures=bool(options.mouse_gestures),
        menu_customizations=bool(options.menu_customizations),
        saved_views=bool(options.saved_views),
    )


def _read_reg_text(file_path: Path) -> tuple[str, str]:
    payload = file_path.read_bytes()
    if payload.startswith(b"\xff\xfe") or payload.startswith(b"\xfe\xff"):
        return payload.decode("utf-16"), "utf-16"
    if payload.startswith(b"\xef\xbb\xbf"):
        return payload.decode("utf-8-sig"), "utf-8-sig"

    for encoding in ("utf-8", "cp1252", "latin-1"):
        try:
            return payload.decode(encoding), encoding
        except Exception:
            continue
    return payload.decode("latin-1", errors="replace"), "latin-1"


def _parse_sldreg_blocks(file_path: Path) -> tuple[list[str], list[tuple[str, list[str]]], str]:
    text, encoding = _read_reg_text(file_path)
    raw_lines = text.splitlines()
    header_lines: list[str] = []
    blocks: list[tuple[str, list[str]]] = []
    current_lines: list[str] = []
    current_key = ""
    in_block = False

    for line in raw_lines:
        stripped = line.strip()
        if stripped.startswith("[") and stripped.endswith("]"):
            if in_block and current_lines:
                blocks.append((current_key, current_lines))
            current_key = stripped[1:-1].strip()
            current_lines = [line]
            in_block = True
            continue

        if in_block:
            current_lines.append(line)
        else:
            header_lines.append(line)

    if in_block and current_lines:
        blocks.append((current_key, current_lines))

    if not header_lines:
        header_lines = ["Windows Registry Editor Version 5.00", ""]

    return header_lines, blocks, encoding


def _is_registry_parent(parent_key: str, child_key: str) -> bool:
    p = parent_key.rstrip("\\").casefold()
    c = child_key.rstrip("\\").casefold()
    return c == p or c.startswith(f"{p}\\")


def _minimize_cleanup_keys(raw_keys: list[str]) -> list[str]:
    unique_keys: list[str] = []
    seen: set[str] = set()
    for key in raw_keys:
        normalized = key.strip()
        if not normalized:
            continue
        low = normalized.casefold()
        if low in seen:
            continue
        seen.add(low)
        unique_keys.append(normalized)

    minimized: list[str] = []
    for key in sorted(unique_keys, key=lambda item: (item.count("\\"), len(item), item.casefold())):
        if any(_is_registry_parent(existing, key) for existing in minimized):
            continue
        minimized.append(key)
    return minimized


def _collect_unsafe_cleanup_parents(excluded_keys: list[str]) -> set[str]:
    unsafe: set[str] = set()
    for key_name in excluded_keys:
        normalized = key_name.strip().rstrip("\\")
        if not normalized:
            continue
        parts = [part for part in normalized.split("\\") if part]
        if not parts:
            continue
        cursor = parts[0]
        unsafe.add(cursor.casefold())
        for part in parts[1:]:
            cursor = f"{cursor}\\{part}"
            unsafe.add(cursor.casefold())
    return unsafe


def _is_recent_registry_key(key_name: str) -> bool:
    low = key_name.casefold()
    tokens = (
        "recent",
        "pinned file list",
    )
    return any(token in low for token in tokens)


def _registry_key_category(key_name: str) -> str:
    low = key_name.casefold()
    if _is_recent_registry_key(key_name):
        return "recent"
    if "\\menu customizations" in low:
        return "menu_customizations"
    if "\\user interface\\saved views" in low:
        return "saved_views"
    if "\\user interface\\settings\\mouse gestures" in low:
        return "mouse_gestures"
    if "\\custom accelerators" in low:
        return "keyboard_shortcuts"
    if "\\user defined macros" in low:
        return "toolbar_macro"
    if "\\user interface\\commandmanager" in low:
        return "toolbar_layout"
    if "\\user interface\\api toolbars" in low:
        return "toolbar_layout"
    if "\\user interface\\toolbars" in low:
        return "toolbar_layout"
    if "\\user interface\\viewtools" in low:
        return "toolbar_layout"
    if "\\simplified interface\\user interface\\viewtools" in low:
        return "toolbar_layout"
    if "\\toolbars" in low:
        return "toolbar_layout"
    return "system_options"


def _should_restore_category(category: str, options: RestoreOptions) -> bool:
    if category == "recent":
        return False
    if category == "system_options":
        return options.system_options
    if category == "toolbar_macro":
        return options.toolbar_layout
    if category == "toolbar_layout":
        return options.toolbar_layout and options.toolbar_mode == "all"
    if category == "keyboard_shortcuts":
        return options.keyboard_shortcuts
    if category == "mouse_gestures":
        return options.mouse_gestures
    if category == "menu_customizations":
        return options.menu_customizations
    if category == "saved_views":
        return options.saved_views
    return False


def _describe_restore_options(options: RestoreOptions) -> str:
    labels: list[str] = []
    if options.system_options:
        labels.append("opzioni sistema")
    if options.toolbar_layout:
        if options.toolbar_mode == "macro_only":
            labels.append("toolbar macro")
        else:
            labels.append("layout toolbar")
    if options.keyboard_shortcuts:
        labels.append("tasti rapidi")
    if options.mouse_gestures:
        labels.append("gesti mouse")
    if options.menu_customizations:
        labels.append("menu")
    if options.saved_views:
        labels.append("viste salvate")
    if not labels:
        return "nessuna opzione"
    return ", ".join(labels)


def write_filtered_sldreg(
    source_file: Path,
    target_file: Path,
    cleanup_before_import: bool = True,
    restore_options: RestoreOptions | None = None,
) -> tuple[bool, str]:
    header_lines, blocks, encoding = _parse_sldreg_blocks(source_file)
    options = _normalize_restore_options(restore_options)
    if not options.has_any_selection():
        return False, "Seleziona almeno una voce da configurare."

    out_lines = list(header_lines)
    if out_lines and out_lines[-1] != "":
        out_lines.append("")

    selected_blocks: list[tuple[str, list[str], str]] = []
    excluded_keys: list[str] = []
    selected_keys: list[str] = []

    for key_name, lines in blocks:
        category = _registry_key_category(key_name)
        if _should_restore_category(category, options):
            selected_blocks.append((key_name, lines, category))
            selected_keys.append(key_name)
        else:
            excluded_keys.append(key_name)

    if not selected_blocks:
        return False, "Nessuna chiave del file .sldreg corrisponde alle opzioni selezionate."

    if cleanup_before_import:
        unsafe_parents = _collect_unsafe_cleanup_parents(excluded_keys)
        cleanup_candidates: list[str] = []
        seen_candidates: set[str] = set()
        for key_name in selected_keys:
            normalized = key_name.strip().rstrip("\\")
            if not normalized:
                continue
            low = normalized.casefold()
            if low in seen_candidates:
                continue
            seen_candidates.add(low)
            if low in unsafe_parents:
                continue
            cleanup_candidates.append(normalized)

        cleanup_keys = _minimize_cleanup_keys(cleanup_candidates)
        for key_name in cleanup_keys:
            out_lines.append(f"[-{key_name}]")
        if cleanup_keys:
            out_lines.append("")

    written_blocks = 0
    for _key_name, lines, _category in selected_blocks:
        out_lines.extend(lines)
        if lines and lines[-1] != "":
            out_lines.append("")
        written_blocks += 1

    if written_blocks == 0:
        return False, "Nessuna chiave valida da importare nel file .sldreg."

    text = "\r\n".join(out_lines).rstrip() + "\r\n"
    target_file.write_text(text, encoding=encoding)
    return True, ""


def import_sldreg_filtered(
    source_file: Path,
    cleanup_before_import: bool = True,
    restore_options: RestoreOptions | None = None,
) -> tuple[bool, str]:
    if not source_file.exists():
        return False, f"File .sldreg non trovato: {source_file}"

    options = _normalize_restore_options(restore_options)
    if not options.has_any_selection():
        return False, "Seleziona almeno una voce da configurare."

    temp_path: Path | None = None
    try:
        fd, temp_name = tempfile.mkstemp(prefix="pdm_swcfg_", suffix=".reg")
        os.close(fd)
        temp_path = Path(temp_name)

        ok_filter, filter_msg = write_filtered_sldreg(
            source_file=source_file,
            target_file=temp_path,
            cleanup_before_import=cleanup_before_import,
            restore_options=options,
        )
        if not ok_filter:
            return False, filter_msg

        completed = subprocess.run(
            ["reg", "import", str(temp_path)],
            capture_output=True,
            text=True,
            check=False,
        )
        if completed.returncode == 0:
            mode = "pulizia + import" if cleanup_before_import else "import senza pulizia"
            scope = _describe_restore_options(options)
            return True, f"Configurazione .sldreg filtrata applicata ({scope}, {mode})."

        error_text = (completed.stderr or completed.stdout or "").strip()
        if not error_text:
            error_text = "Comando reg import terminato con errore."
        return False, error_text
    finally:
        if temp_path is not None:
            try:
                temp_path.unlink(missing_ok=True)
            except Exception:
                pass
