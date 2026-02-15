from __future__ import annotations

from dataclasses import dataclass, asdict, field
from pathlib import Path
from typing import Dict, List, Optional, Literal, Any
import json


Charset = Literal["NUM", "ALPHA", "ALNUM"]
CaseMode = Literal["UPPER", "LOWER"]


@dataclass
class SegmentRule:
    enabled: bool = True
    length: int = 3
    charset: Charset = "ALPHA"
    case: CaseMode = "UPPER"

    def normalize_value(self, value: str) -> str:
        v = (value or "").strip()
        if self.case == "UPPER":
            v = v.upper()
        else:
            v = v.lower()

        if self.charset == "NUM":
            v = "".join(ch for ch in v if ch.isdigit())
        elif self.charset == "ALPHA":
            v = "".join(ch for ch in v if ch.isalpha())
        else:
            v = "".join(ch for ch in v if ch.isalnum())

        # pad/truncate (left pad with 0 only for NUM)
        if self.length > 0:
            if len(v) > self.length:
                v = v[: self.length]
            elif len(v) < self.length:
                pad = "0" if self.charset == "NUM" else "X"
                v = v + (pad * (self.length - len(v)))
        return v


@dataclass
class CodeConfig:
    # Token order fixed for now: MMM sep1 GGGG sep2 0000 [sep3 VVV]
    sep1: str = "_"
    sep2: str = "-"
    sep3: str = "-"
    include_vvv_by_default: bool = False
    vvv_presets: List[str] = field(default_factory=lambda: ["V01", "SKL", "FTM", "FMT"])

    segments: Dict[str, SegmentRule] = field(default_factory=lambda: {
        "MMM": SegmentRule(True, 3, "ALPHA", "UPPER"),
        "GGGG": SegmentRule(True, 4, "ALPHA", "UPPER"),
        "0000": SegmentRule(True, 4, "NUM", "UPPER"),
        "VVV": SegmentRule(True, 3, "ALNUM", "UPPER"),
    })


@dataclass
class SolidWorksConfig:
    archive_root: str = ""              # root archivio
    template_part: str = ""
    template_assembly: str = ""
    template_drawing: str = ""
    property_map: Dict[str, str] = field(default_factory=dict)  # legacy: pdm_field -> SW_PROP
    # Nuovo: lista di mapping, supporta più proprietà SW per lo stesso campo PDM
    property_mappings: List[Dict[str, str]] = field(default_factory=list)  # [{pdm_field, sw_prop}]
    # Descrizione: inserita in PDM alla creazione codice, poi gestita da SolidWorks
    description_prop: str = "DESCRIZIONE"   # nome proprietà SW usata per la descrizione
    # Proprietà custom SolidWorks da leggere (SW → PDM) (oltre la descrizione)
    read_properties: List[str] = field(default_factory=list)



@dataclass
class PDMConfig:
    # Definizione proprietà custom lato PDM (per WORKSPACE)
    # ogni elemento: {name, type, required, default, options}
    custom_properties: List[Dict[str, Any]] = field(default_factory=list)


@dataclass
class BackupConfig:
    enabled: bool = True
    retention_total: int = 30
    daily_enabled: bool = True


@dataclass
class AppConfig:
    code: CodeConfig = field(default_factory=CodeConfig)
    solidworks: SolidWorksConfig = field(default_factory=SolidWorksConfig)
    pdm: PDMConfig = field(default_factory=PDMConfig)
    backup: BackupConfig = field(default_factory=BackupConfig)

    def to_dict(self) -> Dict[str, Any]:
        def convert(obj):
            if hasattr(obj, "__dataclass_fields__"):
                d = asdict(obj)
                return d
            return obj
        return convert(self)

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> "AppConfig":
        def seg_rule(x: Dict[str, Any]) -> SegmentRule:
            return SegmentRule(
                enabled=bool(x.get("enabled", True)),
                length=int(x.get("length", 3)),
                charset=str(x.get("charset", "ALPHA")),
                case=str(x.get("case", "UPPER")),
            )

        code_d = d.get("code", {}) or {}
        segs = {}
        for k, v in (code_d.get("segments", {}) or {}).items():
            if isinstance(v, dict):
                segs[k] = seg_rule(v)

        code = CodeConfig(
            sep1=str(code_d.get("sep1", "_")),
            sep2=str(code_d.get("sep2", "-")),
            sep3=str(code_d.get("sep3", "-")),
            include_vvv_by_default=bool(code_d.get("include_vvv_by_default", False)),
            vvv_presets=list(code_d.get("vvv_presets", ["V01", "SKL", "FTM", "FMT"])),
            segments=segs or CodeConfig().segments,
        )

        sw_d = d.get("solidworks", {}) or {}
        sw = SolidWorksConfig(
            archive_root=str(sw_d.get("archive_root", "")),
            template_part=str(sw_d.get("template_part", "")),
            template_assembly=str(sw_d.get("template_assembly", "")),
            template_drawing=str(sw_d.get("template_drawing", "")),
            property_map=dict(sw_d.get("property_map", {}) or {}),
            property_mappings=list(sw_d.get("property_mappings", []) or []),
            description_prop=str(sw_d.get("description_prop", "DESCRIZIONE")),
            read_properties=list(sw_d.get("read_properties", []) or []),
        )

        pdm_d = d.get("pdm", {}) or {}
        pdm = PDMConfig(
            custom_properties=list(pdm_d.get("custom_properties", []) or [])
        )

        b_d = d.get("backup", {}) or {}
        backup = BackupConfig(
            enabled=bool(b_d.get("enabled", True)),
            retention_total=int(b_d.get("retention_total", 30)),
            daily_enabled=bool(b_d.get("daily_enabled", True)),
        )
        return AppConfig(code=code, solidworks=sw, pdm=pdm, backup=backup)


class ConfigManager:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.cfg = AppConfig()

    def load(self) -> AppConfig:
        if self.path.exists():
            data = json.loads(self.path.read_text(encoding="utf-8"))
            self.cfg = AppConfig.from_dict(data)
        else:
            self.cfg = AppConfig()
            self.save()
        return self.cfg

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(json.dumps(self.cfg.to_dict(), indent=2, ensure_ascii=False), encoding="utf-8")
