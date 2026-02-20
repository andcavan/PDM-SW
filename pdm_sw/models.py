from __future__ import annotations

from dataclasses import dataclass
from typing import Literal


DocType = Literal["PART", "ASSY", "MACHINE", "GROUP"]
State = Literal["WIP", "REL", "IN_REV", "OBS"]


@dataclass
class Document:
    id: int
    code: str
    doc_type: DocType
    mmm: str
    gggg: str
    seq: int
    vvv: str
    revision: int
    state: State
    description: str
    file_wip_path: str
    file_rel_path: str
    file_inrev_path: str

    file_wip_drw_path: str
    file_rel_drw_path: str
    file_inrev_drw_path: str

    created_at: str
    updated_at: str
    obs_prev_state: str = ""
    checked_out: bool = False
    checkout_owner_user: str = ""
    checkout_owner_host: str = ""
    checkout_at: str = ""

    def best_path_for_state(self) -> str:
        if self.state == "WIP":
            return self.file_wip_path
        if self.state == "IN_REV":
            return self.file_inrev_path or self.file_rel_path
        return self.file_rel_path or self.file_wip_path

    def best_drw_path_for_state(self) -> str:
        if self.state == "WIP":
            return self.file_wip_drw_path
        if self.state == "IN_REV":
            return self.file_inrev_drw_path or self.file_rel_drw_path
        return self.file_rel_drw_path or self.file_wip_drw_path

    # Compat alias (vecchie chiamate)
    def best_model_path_for_state(self) -> str:
        return self.best_path_for_state()

    def best_drawing_path_for_state(self) -> str:
        return self.best_drw_path_for_state()
