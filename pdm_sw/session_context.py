from __future__ import annotations

import os
import socket
import uuid
from typing import Dict


def _first_env(keys: list[str]) -> str:
    for k in keys:
        v = str(os.environ.get(k, "") or "").strip()
        if v:
            return v
    return ""


def resolve_session_context(pdm_user_hint: str = "") -> Dict[str, str]:
    """Best-effort identity for audit/lock in shared multi-user setup."""
    pdm_user = (pdm_user_hint or "").strip()
    if not pdm_user:
        pdm_user = _first_env(["PDM_USER", "PDMUSERNAME", "EPDMUSER", "SWPDM_USER"])
    win_user = _first_env(["USERNAME", "USER"])
    domain = _first_env(["USERDOMAIN"])

    user_id = pdm_user or win_user or "unknown"
    display = user_id
    if domain and win_user and user_id == win_user:
        display = f"{domain}\\{win_user}"

    host = socket.gethostname().strip() or "unknown-host"
    session_id = f"{host}-{os.getpid()}-{uuid.uuid4().hex[:8]}"

    source = "PDM" if pdm_user else ("WINDOWS" if win_user else "UNKNOWN")
    return {
        "user_id": user_id,
        "display_name": display,
        "source": source,
        "host": host,
        "session_id": session_id,
    }

