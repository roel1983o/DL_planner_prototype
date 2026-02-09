# Auto-generated runner for DEEL 1 notebook logic
# Executes the cleaned notebook code from a separate file (avoids quoting issues).

from __future__ import annotations

import os
from pathlib import Path
from typing import Dict, Any


HERE = Path(__file__).resolve().parent
CODE_FILE = HERE / "deel1_notebook_code.py"


def run_deel1(
    *,
    templates_path: str,
    beslispad_spread_path: str,
    beslispad_ep_path: str,
    posities_path: str,
    verhalenaanbod_path: str,
    out_path: str,
) -> str:
    ns: Dict[str, Any] = {}

    ns["TEMPLATES_PATH"] = templates_path
    ns["BESLISPAD_SPREAD_PATH"] = beslispad_spread_path
    ns["BESLISPAD_EP_PATH"] = beslispad_ep_path
    ns["POSITIES_PATH"] = posities_path
    ns["VERHALENAANBOD_PATH"] = verhalenaanbod_path
    ns["OUT_PATH"] = out_path

    out_dir = os.path.dirname(out_path)
    os.makedirs(out_dir, exist_ok=True)

    # Run DEEL1 inside the output directory so relative writes land there
    old_cwd = os.getcwd()
    os.chdir(out_dir)
    try:
        code = CODE_FILE.read_text(encoding="utf-8")
        exec(code, ns, ns)
    finally:
        os.chdir(old_cwd)

    # If the notebook wrote with a different name, try to locate the newest xlsx in the folder
    if not os.path.exists(out_path):
        candidates = [
            os.path.join(out_dir, fn)
            for fn in os.listdir(out_dir)
            if fn.lower().endswith(".xlsx")
        ]
        if candidates:
            newest = max(candidates, key=os.path.getmtime)
            if newest != out_path:
                os.replace(newest, out_path)

    if not os.path.exists(out_path):
        raise RuntimeError(f"DEEL 1 finished but output not found: {out_path}")

    return out_path
