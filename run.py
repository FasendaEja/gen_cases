#!/usr/bin/env python
"""Точка входа: python run.py <команда> [параметры]."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from doqa_ai_cases.cli import configure_console, main  # noqa: E402

configure_console()


if __name__ == "__main__":
    raise SystemExit(main())
