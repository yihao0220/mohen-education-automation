from __future__ import annotations

import sys
from typing import TextIO


def _reconfigure_utf8(stream: TextIO | None) -> None:
    reconfigure = getattr(stream, "reconfigure", None)
    if callable(reconfigure):
        reconfigure(encoding="utf-8", errors="strict")


def configure_utf8_stdio() -> None:
    """Keep Chinese CLI output readable when Windows redirects stdout/stderr."""

    _reconfigure_utf8(sys.stdout)
    _reconfigure_utf8(sys.stderr)
