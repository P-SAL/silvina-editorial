"""
conftest.py
Pytest configuration — applies to all tests.
"""
import sys
from unittest.mock import MagicMock

# Reconfigure stdout/stderr to UTF-8 so that emoji in source print() calls
# don't cause UnicodeEncodeError on Windows cp1252 consoles.
if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
if hasattr(sys.stderr, 'reconfigure'):
    try:
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# Inject win32com mocks at the start of the pytest session.
# This ensures modules that import win32com at the top level do not fail.
if 'win32com' not in sys.modules:
    _win32com_client = MagicMock()
    _win32com = MagicMock()
    _win32com.client = _win32com_client
    sys.modules['win32com'] = _win32com
    sys.modules['win32com.client'] = _win32com_client
    sys.modules['pythoncom'] = MagicMock()
