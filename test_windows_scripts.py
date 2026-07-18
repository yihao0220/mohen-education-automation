from __future__ import annotations

import codecs
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
WINDOWS_SCRIPTS = (
    PROJECT_ROOT / "scripts" / "setup_windows.ps1",
    PROJECT_ROOT / "scripts" / "verify_windows.ps1",
)


def test_windows_powershell_scripts_use_utf8_bom() -> None:
    for script in WINDOWS_SCRIPTS:
        payload = script.read_bytes()
        assert payload.startswith(codecs.BOM_UTF8), (
            f"{script.name} 必须带 UTF-8 BOM，兼容 Windows PowerShell 5.1 的中文源码解析"
        )
        payload.decode("utf-8-sig")
