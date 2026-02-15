from __future__ import annotations

from dataclasses import dataclass
from .sw_api import get_solidworks_app, _get_or_call


@dataclass
class SWStatus:
    ok: bool
    message: str
    version: str = ""
    details: str = ""


def test_solidworks_connection() -> SWStatus:
    sw, res = get_solidworks_app(visible=False, timeout_s=20.0)
    if not res.ok or sw is None:
        return SWStatus(False, res.message, "", res.details)

    ver = ""
    try:
        ver = str(_get_or_call(sw, "RevisionNumber"))
    except Exception:
        try:
            ver = str(_get_or_call(sw, "GetCurrentVersion"))
        except Exception:
            ver = ""
    return SWStatus(True, res.message, ver, "")
