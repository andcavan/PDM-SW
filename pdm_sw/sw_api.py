from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Optional, Tuple, Dict
import time
import struct
import ctypes
import sys
import os
from pathlib import Path

def _norm_path(p: str) -> str:
    """Normalizza un path Windows per SolidWorks (backslash + normpath)."""
    try:
        p = p.strip().strip('"')
    except Exception:
        pass
    p = p.replace('/', '\\')
    try:
        p = os.path.normpath(p)
    except Exception:
        pass
    return p


def _byref_i4(value: int = 0):
    """Crea un VARIANT byref int32 per chiamate COM (OpenDoc6/Save3).

    In late-binding pywin32, passare interi semplici per parametri 'ByRef' può dare:
    (-2147352561, 'Parametro non facoltativo.', None, None)
    """
    try:
        import pythoncom  # type: ignore
        from win32com.client import VARIANT  # type: ignore
        return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, int(value))
    except Exception:
        # Se non disponibile, restituiamo un int (potrebbe non funzionare su chiamate ByRef).
        return int(value)


def _i4_value(v: Any) -> int:
    """Estrae un int da VARIANT byref (o da valore normale)."""
    try:
        return int(getattr(v, "value"))
    except Exception:
        pass
    try:
        return int(v)
    except Exception:
        return 0


def _byref_bstr(value: str = ""):
    """VARIANT byref BSTR (stringa) per chiamate COM (Get4)."""
    try:
        import pythoncom  # type: ignore
        from win32com.client import VARIANT  # type: ignore
        return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, str(value))
    except Exception:
        return str(value)


def _byref_bool(value: bool = False):
    """VARIANT byref BOOL per chiamate COM."""
    try:
        import pythoncom  # type: ignore
        from win32com.client import VARIANT  # type: ignore
        return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BOOL, bool(value))
    except Exception:
        return bool(value)




# HRESULTs
_RPC_FAILED_HRESULT = -2147023170
_RPC_E_CALL_REJECTED = -2147418111
_RPC_E_SERVERCALL_RETRYLATER = -2147417846

_OLE_FILTER = None


@dataclass
class SWResult:
    ok: bool
    message: str
    details: str = ""



# SolidWorks constants (swconst):
# swCustomInfoType_e.swCustomInfoText = 30
# swCustomPropertyAddOption_e.swCustomPropertyReplaceValue = 2

def set_custom_properties(model_doc: Any, props: Dict[str, str]) -> None:
    """Set (or replace) custom properties on the active configuration (general properties)."""
    if not props:
        return
    try:
        ext = model_doc.Extension
        mgr = ext.CustomPropertyManager("")
    except Exception:
        return
    for name, value in props.items():
        n = str(name).strip()
        if not n:
            continue
        v = "" if value is None else str(value)
        try:
            # Prefer Set2 when available
            if hasattr(mgr, "Set2"):
                mgr.Set2(n, v)
            else:
                # Add3(FieldName, FieldType, FieldValue, AddOption)
                mgr.Add3(n, 30, v, 2)
        except Exception:
            try:
                mgr.Add3(n, 30, v, 2)
            except Exception:
                pass


def get_custom_properties(model_doc: Any) -> Dict[str, str]:
    """Legge le proprietà custom (generali) e ritorna dict nome->valore RISOLTO.

    Alcune proprietà in SolidWorks possono essere espressioni/collegamenti (es. $PRP:..., SW-Material@...).
    Per ottenere il valore risolto usiamo CustomPropertyManager.Get4 con parametri ByRef corretti.
    Fallback su ModelDoc2.CustomInfoValue2/CustomInfo2.
    """
    out: Dict[str, str] = {}
    if model_doc is None:
        return out

    mgr = None
    names: list[str] = []
    try:
        ext = model_doc.Extension
        mgr = ext.CustomPropertyManager("")
        try:
            names = list(mgr.GetNames() or [])
        except Exception:
            names = []
    except Exception:
        mgr = None
        names = []

    if not names:
        try:
            if hasattr(model_doc, "GetCustomInfoNames2"):
                names = list(model_doc.GetCustomInfoNames2("") or [])
        except Exception:
            names = []

    for n in names:
        try:
            key = str(n).strip()
            if not key:
                continue

            val_s = ""

            # 1) Get4 -> raw + resolved
            if mgr is not None and hasattr(mgr, "Get4"):
                try:
                    v_out = _byref_bstr("")
                    r_out = _byref_bstr("")
                    mgr.Get4(key, False, v_out, r_out)
                    raw = getattr(v_out, "value", v_out)
                    res = getattr(r_out, "value", r_out)
                    val_s = str(res or raw or "")
                except Exception:
                    val_s = ""

            # 2) fallback (può restituire formule)
            if not val_s and hasattr(model_doc, "CustomInfoValue2"):
                try:
                    val_s = str(model_doc.CustomInfoValue2("", key) or "")
                except Exception:
                    val_s = ""
            if not val_s and hasattr(model_doc, "CustomInfo2"):
                try:
                    val_s = str(model_doc.CustomInfo2("", key) or "")
                except Exception:
                    val_s = ""

            # rimuovi doppi apici esterni
            if len(val_s) >= 2 and val_s.startswith('"') and val_s.endswith('"'):
                val_s = val_s[1:-1]

            out[key] = val_s
        except Exception:
            continue

    return out

    names: list[str] = []
    try:
        ext = model_doc.Extension
        mgr = ext.CustomPropertyManager("")
        try:
            names = list(mgr.GetNames() or [])
        except Exception:
            names = []
    except Exception:
        names = []

    # Fallback: prova a recuperare i nomi dal ModelDoc2 (se disponibili)
    if not names:
        try:
            if hasattr(model_doc, "GetCustomInfoNames2"):
                names = list(model_doc.GetCustomInfoNames2("") or [])
        except Exception:
            names = []

    for n in names:
        try:
            key = str(n).strip()
            if not key:
                continue
            val = ""
            # resolved value (se disponibile)
            if hasattr(model_doc, "CustomInfoValue2"):
                try:
                    val = model_doc.CustomInfoValue2("", key)
                except Exception:
                    val = ""
            if not val and hasattr(model_doc, "CustomInfo2"):
                try:
                    val = model_doc.CustomInfo2("", key)
                except Exception:
                    val = ""
            out[key] = "" if val is None else str(val)
        except Exception:
            continue
    return out




def _coinit() -> None:
    try:
        import pythoncom  # type: ignore
        pythoncom.CoInitialize()
    except Exception:
        pass


def _register_ole_message_filter() -> None:
    """Mitiga 'Call rejected by callee' quando SolidWorks è occupato."""
    global _OLE_FILTER
    try:
        import pythoncom  # type: ignore
    except Exception:
        return

    class OleMessageFilter:
        def HandleInComingCall(self, *args):
            return 0  # SERVERCALL_ISHANDLED

        def RetryRejectedCall(self, *args):
            # args[2] = rejectType (2 = retry later)
            try:
                reject_type = args[2]
                if reject_type == 2:
                    return 250  # retry in 250ms
            except Exception:
                pass
            return -1  # cancel

        def MessagePending(self, *args):
            return 2  # PENDINGMSG_WAITDEFPROCESS

    try:
        _OLE_FILTER = OleMessageFilter()
        pythoncom.CoRegisterMessageFilter(_OLE_FILTER)
    except Exception:
        pass


def _get_or_call(obj: Any, attr_name: str):
    a = getattr(obj, attr_name, None)
    if a is None:
        raise AttributeError(attr_name)
    return a() if callable(a) else a


def _is_transient(exc: Exception) -> bool:
    s = str(exc).lower()
    return ("rejected" in s) or ("retry" in s) or (str(_RPC_E_CALL_REJECTED) in s) or (str(_RPC_E_SERVERCALL_RETRYLATER) in s)


def _sw_ping(sw: Any) -> bool:
    for name in ("CommandInProgress", "Visible", "RevisionNumber", "GetProcessID"):
        try:
            _ = _get_or_call(sw, name)
            return True
        except Exception:
            continue
    return False


def _rot_enum_sldworks_apps() -> list[Any]:
    """Ritorna una lista di istanze SolidWorks trovate nella ROT (Running Object Table).
    Utile quando esistono più istanze o GetActiveObject non punta a quella corretta.
    """
    apps: list[Any] = []
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception:
        return apps

    try:
        pythoncom.CoInitialize()
        rot = pythoncom.GetRunningObjectTable()
        enum = rot.EnumRunning()
        ctx = pythoncom.CreateBindCtx(0)
        while True:
            monikers = enum.Next(1)
            if not monikers:
                break
            mon = monikers[0]
            try:
                name = mon.GetDisplayName(ctx, None)
            except Exception:
                continue
            if "SldWorks.Application" not in str(name):
                continue
            try:
                obj = rot.GetObject(mon)
                sw = win32com.client.Dispatch(obj)
                apps.append(sw)
            except Exception:
                continue
    except Exception:
        return apps

    # de-dup per PID (se disponibile)
    uniq: dict[int, Any] = {}
    for sw in apps:
        try:
            pid = _get_or_call(sw, "GetProcessID")
        except Exception:
            pid = id(sw)
        try:
            pid_int = int(pid)
        except Exception:
            pid_int = id(sw)
        uniq[pid_int] = sw
    return list(uniq.values())


def _active_doc_path(sw: Any) -> str:
    """Prova a leggere il path del documento attivo da un'istanza SolidWorks."""
    try:
        doc = getattr(sw, "IActiveDoc2", None)
        doc = doc() if callable(doc) else doc
        if not doc:
            doc = getattr(sw, "ActiveDoc", None)
            doc = doc() if callable(doc) else doc
        if not doc:
            return ""
        gp = getattr(doc, "GetPathName", None)
        v = gp() if callable(gp) else gp
        return str(v or "")
    except Exception:
        return ""


def _select_best_sw_app(candidates: list[Any], prefer_pid: Optional[int] = None, prefer_doc_path: Optional[str] = None) -> Optional[Any]:
    """Seleziona l'istanza SolidWorks più adatta tra quelle candidate."""
    if not candidates:
        return None

    # 1) match PID
    if prefer_pid:
        for sw in candidates:
            try:
                pid = _get_or_call(sw, "GetProcessID")
                if int(pid) == int(prefer_pid):
                    return sw
            except Exception:
                continue

    # 2) match active doc path
    if prefer_doc_path:
        p_norm = str(prefer_doc_path).replace("/", "\\").lower()
        # Se il documento è aperto in una specifica istanza, preferiscila
        for sw in candidates:
            try:
                if hasattr(sw, "GetOpenDocumentByName"):
                    d = sw.GetOpenDocumentByName(str(prefer_doc_path))
                    if d is not None:
                        return sw
            except Exception:
                pass

        # Match sul documento attivo
        for sw in candidates:
            try:
                ap = _active_doc_path(sw).replace("/", "\\").lower()
                if ap and ap == p_norm:
                    return sw
            except Exception:
                continue

    # 3) preferisci una istanza con doc attivo
    for sw in candidates:
        try:
            if _active_doc_path(sw):
                return sw
        except Exception:
            continue

    # 4) fallback: prima
    return candidates[0]


def get_solidworks_app(
    visible: bool = False,
    timeout_s: float = 30.0,
    allow_launch: bool = True,
    prefer_pid: Optional[int] = None,
    prefer_doc_path: Optional[str] = None,
) -> Tuple[Optional[Any], SWResult]:
    """Ritorna (swApp, SWResult).

    allow_launch=False: NON avvia una nuova istanza (utile in macro SolidWorks).
    prefer_pid / prefer_doc_path: aiuta a scegliere l'istanza corretta se ne esistono più.
    """
    _coinit()
    _register_ole_message_filter()

    try:
        import win32com.client  # type: ignore
    except Exception as e:
        return None, SWResult(False, "pywin32 non installato.", f"Dettagli: {e}")

    sw = None

    # Prova ROT enumeration prima (più affidabile in scenari multi-istanza)
    try:
        rot_candidates = _rot_enum_sldworks_apps()
        sw = _select_best_sw_app(rot_candidates, prefer_pid=prefer_pid, prefer_doc_path=prefer_doc_path)
    except Exception:
        sw = None

    # Fallback: GetActiveObject
    if sw is None:
        try:
            sw = win32com.client.GetActiveObject("SldWorks.Application")
        except Exception:
            sw = None

    # Se ancora None, prova Dispatch (non DispatchEx) se consentito
    if sw is None and allow_launch:
        try:
            sw = win32com.client.Dispatch("SldWorks.Application")
        except Exception:
            sw = None

    # Ultimo fallback: DispatchEx (nuova istanza) solo se consentito
    if sw is None and allow_launch:
        try:
            sw = win32com.client.DispatchEx("SldWorks.Application")
        except Exception as e:
            return None, SWResult(False, "Impossibile avviare SolidWorks via COM.", str(e))

    if sw is None:
        return None, SWResult(False, "SolidWorks non trovato (COM).", "Avvia SolidWorks e riprova.")

    try:
        sw.Visible = bool(visible)
    except Exception:
        pass

    # attesa reattività
    t0 = time.time()
    last_exc = None
    while time.time() - t0 < timeout_s:
        try:
            if _sw_ping(sw):
                return sw, SWResult(True, "Connesso a SolidWorks.")
        except Exception as e:
            last_exc = e
        time.sleep(0.2)

    return sw, SWResult(False, "SolidWorks non risponde (COM).", str(last_exc) if last_exc else None)


def create_new_from_template(sw: Any, template_path: str) -> Any:
    """Crea un nuovo documento da template."""
    # SolidWorks API: NewDocument(template, paperSize, width, height)
    template_path = _norm_path(template_path)
    return sw.NewDocument(template_path, 0, 0.0, 0.0)




def save_as_doc(doc, out_path: str) -> SWResult:
    """SaveAs sul documento già aperto (ModelDoc2), con fallback robusti.

    Nota: in alcune versioni/contesti COM SaveAs può restituire False anche
    quando il file viene comunque creato. In quel caso consideriamo successo
    se il target esiste ed è stato aggiornato.
    """
    try:
        out_path = _norm_path(str(out_path))
        dst = Path(out_path)
        ensure_parent_ok = True
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
        except Exception:
            ensure_parent_ok = False

        prev_mtime_ns = None
        try:
            if dst.exists():
                prev_mtime_ns = dst.stat().st_mtime_ns
        except Exception:
            prev_mtime_ns = None

        attempts = []

        # 1) IModelDocExtension.SaveAs: generalmente la via più affidabile.
        try:
            ext = getattr(doc, "Extension", None)
            if ext is not None and hasattr(ext, "SaveAs"):
                err = _byref_i4(0)
                warn = _byref_i4(0)
                ok = bool(ext.SaveAs(out_path, 0, 1, None, err, warn))
                ecode = _i4_value(err)
                wcode = _i4_value(warn)
                attempts.append(f"Extension.SaveAs ok={ok} err={ecode} warn={wcode}")
                if ok:
                    return SWResult(True, f"Salvato come: {out_path}", f"err={ecode}; warn={wcode}")
        except Exception as e:
            attempts.append(f"Extension.SaveAs EXC={e}")

        # 2) SaveAs3 con opzione silent.
        try:
            ok = bool(doc.SaveAs3(out_path, 0, 1))
            attempts.append(f"SaveAs3(0,1) ok={ok}")
            if ok:
                return SWResult(True, f"Salvato come: {out_path}")
        except Exception as e:
            attempts.append(f"SaveAs3(0,1) EXC={e}")

        # 3) SaveAs3 legacy senza opzioni.
        try:
            ok = bool(doc.SaveAs3(out_path, 0, 0))
            attempts.append(f"SaveAs3(0,0) ok={ok}")
            if ok:
                return SWResult(True, f"Salvato come: {out_path}")
        except Exception as e:
            attempts.append(f"SaveAs3(0,0) EXC={e}")

        # 4) Fallback SaveAs (alcune installazioni espongono questa firma).
        try:
            if hasattr(doc, "SaveAs"):
                ok = bool(doc.SaveAs(out_path))
                attempts.append(f"SaveAs ok={ok}")
                if ok:
                    return SWResult(True, f"Salvato come: {out_path}")
        except Exception as e:
            attempts.append(f"SaveAs EXC={e}")

        # Fallback pragmatico: file creato/aggiornato anche se API segnala False.
        try:
            if dst.exists():
                st = dst.stat()
                changed = prev_mtime_ns is None or st.st_mtime_ns != prev_mtime_ns
                if st.st_size > 0 and changed:
                    det = " | ".join(attempts)
                    return SWResult(
                        True,
                        f"Salvato come: {out_path} (con warning API)",
                        det,
                    )
        except Exception:
            pass

        det = " | ".join(attempts)
        if not ensure_parent_ok:
            det = ("mkdir parent fallita | " + det).strip()
        return SWResult(False, f"SaveAs fallito: {out_path}", det)
    except Exception as e:
        return SWResult(False, f"SaveAs fallito: {e}")

def save_doc(doc: Any, out_path: str) -> None:
    # SaveAs3(FileName, SaveAsOptions, Errors)
    # use 0 options; errors byref not handled -> returns bool usually
    doc.SaveAs3(out_path, 0, 0)


def open_doc(sw: Any, file_path: str, silent: bool = True) -> Any:
    """Apre un documento esistente in SolidWorks (best-effort).

    Usa sempre parametri ByRef corretti per OpenDoc6 (errors/warnings).
    """
    file_path = _norm_path(file_path)
    ext = Path(file_path).suffix.lower()
    # swDocumentTypes_e: part=1, assembly=2, drawing=3
    if ext == ".sldprt":
        doc_type = 1
    elif ext == ".sldasm":
        doc_type = 2
    elif ext == ".slddrw":
        doc_type = 3
    else:
        doc_type = 0
    # swOpenDocOptions_Silent = 1
    options = 1 if silent else 0

    # Se già aperto, restituisci l'istanza
    try:
        if hasattr(sw, "GetOpenDocumentByName"):
            already = sw.GetOpenDocumentByName(file_path)
            if already is not None:
                return already
    except Exception:
        pass

    err = _byref_i4(0)
    warn = _byref_i4(0)

    try:
        return sw.OpenDoc6(file_path, doc_type, options, "", err, warn)
    except Exception as e:
        # Non provare firme 'accorciate' (generano spesso "Parametro non facoltativo")
        raise RuntimeError(f"OpenDoc6 fallita: {e}")


def save_existing_doc(doc: Any) -> None:
    """Salva un documento già aperto (best-effort).

    Save3 richiede errors/warnings ByRef: usa _byref_i4 per evitare 'Parametro non facoltativo'.
    """
    if doc is None:
        return
    try:
        if hasattr(doc, "Save3"):
            err = _byref_i4(0)
            warn = _byref_i4(0)
            doc.Save3(1, err, warn)
            return
    except Exception:
        pass
    try:
        if hasattr(doc, "Save"):
            doc.Save()
    except Exception:
        pass


def create_model_file(sw: Any, template_path: str, out_path: str, props: Dict[str, str] | None = None) -> SWResult:
    try:
        template_path_n = _norm_path(template_path)
        if not Path(template_path_n).is_file():
            return SWResult(
                False,
                "Template non trovato o non accessibile.",
                f"MODEL template: {template_path_n}\nSuggerimento: usa percorso locale oppure UNC (\\\\server\\share\\...) e backslash.",
            )
        doc = create_new_from_template(sw, template_path_n)
        if doc is None:
            return SWResult(
                False,
                "NewDocument ha restituito None (template non valido o non leggibile da SolidWorks).",
                f"MODEL template: {template_path_n}\nVerifica che il file sia un *template* valido (.prtdot/.asmdot) e che SolidWorks possa accedere al percorso (drive mappato/permessi).",
            )
        if props:
            set_custom_properties(doc, props)
        save_doc(doc, out_path)
        return SWResult(True, "File modello creato.")
    except Exception as e:
        return SWResult(False, "Creazione file modello fallita.", str(e))


def create_drawing_file(sw: Any, template_path: str, out_path: str, props: Dict[str, str] | None = None) -> SWResult:
    try:
        template_path_n = _norm_path(template_path)
        if not Path(template_path_n).is_file():
            return SWResult(
                False,
                "Template non trovato o non accessibile.",
                f"DRW template: {template_path_n}\nSuggerimento: usa percorso locale oppure UNC (\\\\server\\share\\...) e backslash.",
            )
        doc = create_new_from_template(sw, template_path_n)
        if doc is None:
            return SWResult(
                False,
                "NewDocument ha restituito None (template non valido o non leggibile da SolidWorks).",
                f"DRW template: {template_path_n}\nVerifica che il file sia un *template* valido (.drwdot) e che SolidWorks possa accedere al percorso (drive mappato/permessi).",
            )
        if props:
            set_custom_properties(doc, props)
        save_doc(doc, out_path)
        return SWResult(True, "File disegno creato.")
    except Exception as e:
        return SWResult(False, "Creazione file disegno fallita.", str(e))


def close_doc(sw: Any, doc: Any = None, file_path: str | None = None) -> None:
    """Chiude un documento aperto in SolidWorks (best-effort)."""
    if sw is None:
        return
    name = None
    try:
        if doc is not None and hasattr(doc, "GetTitle"):
            name = doc.GetTitle()
    except Exception:
        name = None
    if not name and file_path:
        try:
            # SolidWorks spesso vuole il titolo (nome file), non necessariamente il path
            name = Path(_norm_path(file_path)).name
        except Exception:
            name = file_path
    if not name:
        return
    try:
        sw.CloseDoc(name)
    except Exception:
        # ignora
        pass
