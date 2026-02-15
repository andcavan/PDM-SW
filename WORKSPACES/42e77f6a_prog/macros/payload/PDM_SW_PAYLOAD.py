# -*- coding: utf-8 -*-
# PDM SolidWorks Payload (launcher)
# Generato dal PDM - workspace: 42e77f6a
#
# Questo file lancia la UI minimale per Codifica/Workflow da SolidWorks.

import sys
from pathlib import Path

def _extract_pdm_root(argv):
    try:
        i = argv.index("--pdm-root")
        return argv[i + 1]
    except Exception:
        return ""

root = _extract_pdm_root(sys.argv)
if root:
    pdm_root = Path(root)
    if str(pdm_root) not in sys.path:
        sys.path.insert(0, str(pdm_root))

from pdm_sw.macro_runtime import main  # noqa

if __name__ == "__main__":
    main()
