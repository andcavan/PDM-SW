from __future__ import annotations

import struct
import sys
import time
import ctypes

try:
    import pythoncom  # type: ignore
    import pywintypes  # type: ignore
    import win32com.client  # type: ignore
except Exception as e:
    print("Dipendenze COM mancanti:", e)
    raise SystemExit(2)


def _get_or_call(obj, name: str):
    a = getattr(obj, name, None)
    if a is None:
        raise AttributeError(name)
    return a() if callable(a) else a


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def main() -> int:
    pythoncom.CoInitialize()
    print("Python:", sys.version.replace("\n", " "))
    print("Bitness:", struct.calcsize("P") * 8)
    print("Admin:", is_admin())
    print()

    try:
        sw = win32com.client.GetActiveObject("SldWorks.Application")
        print("GetActiveObject: OK")
    except Exception as e:
        print("GetActiveObject: FAIL ->", e)
        return 1

    tests = [
        ("RevisionNumber", lambda: _get_or_call(sw, "RevisionNumber")),
        ("GetCurrentVersion", lambda: _get_or_call(sw, "GetCurrentVersion")),
        ("GetProcessID", lambda: _get_or_call(sw, "GetProcessID") if hasattr(sw, "GetProcessID") else "n/a"),
        ("Visible", lambda: getattr(sw, "Visible", "n/a")),
        ("CommandInProgress", lambda: getattr(sw, "CommandInProgress", "n/a")),
        ("ActiveDoc", lambda: sw.ActiveDoc),
    ]

    for name, fn in tests:
        try:
            t0 = time.time()
            val = fn()
            dt = (time.time() - t0) * 1000
            print(f"{name:18s} OK ({dt:,.0f} ms): {val}")
        except Exception as e:
            if isinstance(e, pywintypes.com_error):
                print(f"{name:18s} COM_ERROR hresult={e.hresult}: {e}")
            else:
                print(f"{name:18s} ERROR {type(e).__name__}: {e}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
