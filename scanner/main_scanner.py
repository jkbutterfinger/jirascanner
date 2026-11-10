#!/usr/bin/env python3
"""Kompatibilitäts-Wrapper für den historischen Scanner-Einstiegspunkt.

Die kanonische Implementierung liegt in ``main.py``. Dieser Wrapper erlaubt
weiterhin Aufrufe wie ``python -m scanner.main_scanner``.
"""

from __future__ import annotations

import sys
from typing import List, Optional

from main import main as _main


def main(argv: Optional[List[str]] = None) -> int:
    """Delegiert an den kanonischen CLI-Einstiegspunkt."""
    return _main(argv)


if __name__ == "__main__":
    sys.exit(main())
