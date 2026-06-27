"""
conftest.py  —  Mock heavy dependencies before any module import.

Stubs out yfinance, openpyxl, and akshare so the scoring modules can be
imported without a live broker connection or Excel files on disk.
numpy / pandas are kept real because the scoring functions use them.
"""

import sys
import os
import logging
from unittest.mock import MagicMock, patch

# ── 1. Stub openpyxl (styles / utils used at module level) ────────────────
_styles = MagicMock(name="openpyxl.styles")
_styles.PatternFill = MagicMock(side_effect=lambda *a, **kw: MagicMock())
_styles.Alignment   = MagicMock(side_effect=lambda *a, **kw: MagicMock())
_styles.Font        = MagicMock(side_effect=lambda *a, **kw: MagicMock())
_styles.Border      = MagicMock(side_effect=lambda *a, **kw: MagicMock())
_styles.Side        = MagicMock(side_effect=lambda *a, **kw: MagicMock())

_utils = MagicMock(name="openpyxl.utils")
_utils.get_column_letter = lambda n: chr(64 + n) if 1 <= n <= 26 else "Z"

_openpyxl = MagicMock(name="openpyxl")
_openpyxl.styles = _styles
_openpyxl.utils  = _utils
_openpyxl.load_workbook = MagicMock(return_value=MagicMock())

_chart = MagicMock(name="openpyxl.chart")
_chart.LineChart  = MagicMock(return_value=MagicMock())
_chart.Reference  = MagicMock(return_value=MagicMock())
_openpyxl.chart   = _chart

sys.modules.setdefault("openpyxl",               _openpyxl)
sys.modules.setdefault("openpyxl.styles",         _styles)
sys.modules.setdefault("openpyxl.utils",          _utils)
sys.modules.setdefault("openpyxl.chart",          _chart)

# ── 2. Stub yfinance ───────────────────────────────────────────────────────
_yf = MagicMock(name="yfinance")
sys.modules.setdefault("yfinance", _yf)

# ── 3. Stub akshare (optional import, try/except in modules) ──────────────
sys.modules.setdefault("akshare", MagicMock(name="akshare"))

# ── 4. Redirect log-file creation to a no-op ──────────────────────────────
# daily_importer creates a FileHandler at import time; redirect to NullHandler
_orig_file_handler = logging.FileHandler

class _NullFileHandler(logging.NullHandler):
    def __init__(self, *a, **kw):
        super().__init__()
    def emit(self, record):
        pass

logging.FileHandler = _NullFileHandler

# ── 5. Add project root to sys.path ───────────────────────────────────────
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)
