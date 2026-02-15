from __future__ import annotations

import struct
import sys
import time
import ctypes

try:
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
    import pywintypes  # type: ignore
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

    import subprocess
    try:
        out = subprocess.check_output(["tasklist", "/FI", "IMAGENAME eq SLDWORKS.exe"], text=True, errors="ignore")
        print(out.strip())
    except Exception:
        pass
    print()

    try:
        sw = win32com.client.GetActiveObject("SldWorks.Application")
        print("GetActiveObject: OK (istanza già aperta)")
    except Exception as e:
        print("GetActiveObject: FAIL ->", e)
        try:
            sw = win32com.client.DispatchEx("SldWorks.Application")
            print("DispatchEx: OK (nuova istanza avviata)")
        except Exception as e2:
            print("DispatchEx: FAIL ->", e2)
            return 1

    print("Attendo che SolidWorks risponda...")
    last_err = None
    for i in range(180):  # 90s
        v = None
        try:
            try:
                v = _get_or_call(sw, "RevisionNumber")
            except Exception as e1:
                last_err = e1
                try:
                    v = _get_or_call(sw, "GetCurrentVersion")
                except Exception as e2:
                    last_err = e2
                    try:
                        v = f"PID={_get_or_call(sw, 'GetProcessID')}"
                    except Exception as e3:
                        last_err = e3
            if v is not None:
                print("Risponde. Info:", v)
                return 0
        except Exception as e:
            last_err = e

        if last_err and (i < 5 or i % 20 == 0):
            if isinstance(last_err, pywintypes.com_error):
                print(f"[{i*0.5:5.1f}s] COM error hresult={last_err.hresult}: {last_err}")
            else:
                print(f"[{i*0.5:5.1f}s] {type(last_err).__name__}: {last_err}")

        time.sleep(0.5)

    print("NON risponde dopo 90s.")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
