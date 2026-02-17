from __future__ import annotations

from typing import Optional
from .config import AppConfig
from .models import DocType


def format_seq(seq: int, length: int = 4) -> str:
    return str(seq).zfill(length)


def build_code(cfg: AppConfig, mmm: str, gggg: str, seq: int, vvv: str = "", force_vvv: Optional[bool] = None, doc_type: Optional[DocType] = None) -> str:
    """Genera codice per PART/ASSY standard."""
    segs = cfg.code.segments
    mmm_v = segs["MMM"].normalize_value(mmm)
    gggg_v = segs["GGGG"].normalize_value(gggg)
    seq_v = str(seq).zfill(segs["0000"].length)

    include_vvv = cfg.code.include_vvv_by_default if force_vvv is None else bool(force_vvv)
    if include_vvv and vvv:
        vvv_v = segs["VVV"].normalize_value(vvv)
        # Nuovo formato: [MMM]_[GGGG]-[VVV]-[0000]
        return f"{mmm_v}{cfg.code.sep1}{gggg_v}{cfg.code.sep2}{vvv_v}{cfg.code.sep3}{seq_v}"
    # Senza variante: [MMM]_[GGGG]-[0000]
    return f"{mmm_v}{cfg.code.sep1}{gggg_v}{cfg.code.sep2}{seq_v}"


def build_machine_code(cfg: AppConfig, mmm: str, ver_seq: int) -> str:
    """Genera codice per MACHINE: MMM-V####."""
    segs = cfg.code.segments
    mmm_v = segs["MMM"].normalize_value(mmm)
    vnum_v = str(ver_seq).zfill(segs.get("VNUM", segs["0000"]).length)
    return f"{mmm_v}{cfg.code.sep2}V{vnum_v}"


def build_group_code(cfg: AppConfig, mmm: str, gggg: str, ver_seq: int) -> str:
    """Genera codice per GROUP: MMM_GGGG-V####."""
    segs = cfg.code.segments
    mmm_v = segs["MMM"].normalize_value(mmm)
    gggg_v = segs["GGGG"].normalize_value(gggg)
    vnum_v = str(ver_seq).zfill(segs.get("VNUM", segs["0000"]).length)
    return f"{mmm_v}{cfg.code.sep1}{gggg_v}{cfg.code.sep2}V{vnum_v}"
