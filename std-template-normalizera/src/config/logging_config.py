"""
Centralized logging for std-template-normalizer.
Writes to the same log file as the C# app: %APPDATA%\\ste_tool_studio\\ste_tool_studio.log
Use get_logger(__name__) in modules.
"""
import logging
import os
import sys

from .constants import APP_DATA_FOLDER_NAME

# Same file name as C# AppConstants.LogFileName so both projects use one log file
LOG_FILE_NAME = "ste_tool_studio.log"


def _log_file_path():
    appdata = os.getenv("APPDATA") or os.path.expanduser("~\\AppData\\Roaming")
    return os.path.join(appdata, APP_DATA_FOLDER_NAME, LOG_FILE_NAME)


def _ensure_log_dir():
    path = _log_file_path()
    d = os.path.dirname(path)
    os.makedirs(d, exist_ok=True)
    return path


def setup_logging(level=logging.DEBUG):
    """
    Configure root logger to write to shared log file (and stderr for debugging).
    Idempotent; safe to call multiple times.
    """
    root = logging.getLogger()
    if root.handlers:
        return root

    root.setLevel(level)
    fmt = logging.Formatter(
        "[%(asctime)s] %(levelname)-8s [%(name)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    try:
        file_path = _ensure_log_dir()
        fh = logging.FileHandler(file_path, encoding="utf-8", mode="a")
        fh.setLevel(level)
        fh.setFormatter(fmt)
        root.addHandler(fh)
    except Exception as e:
        # Fallback: log to stderr only
        sys.stderr.write(f"Could not create log file: {e}\n")

    ch = logging.StreamHandler(sys.stderr)
    ch.setLevel(level)
    ch.setFormatter(fmt)
    root.addHandler(ch)

    return root


def get_logger(name: str):
    """Return a logger for the given module name. Calls setup_logging() on first use."""
    setup_logging()
    return logging.getLogger(name)
