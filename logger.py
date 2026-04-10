"""
logger.py — Centralized logging for Hamdaz App
================================================
Usage:
    from logger import log, DEBUG_MODE

Toggle debug output:
    - Set env var:  DEBUG=true   (or false)
    - Or flip the fallback default below.

Log levels used:
    log.info(msg)    — normal operational info  (always shown)
    log.warn(msg)    — non-fatal warnings       (always shown)
    log.error(msg)   — errors / exceptions      (always shown)
    log.debug(msg)   — verbose developer detail (only when DEBUG=true)
    log.success(msg) — positive confirmations   (only when DEBUG=true)
"""

import logging
import os
import sys
from datetime import datetime

# ──────────────────────────────────────────────
# Toggle: read from env var, default to False
# Set DEBUG=true in your .env to enable verbose logs
# ──────────────────────────────────────────────
DEBUG_MODE: bool = os.getenv("DEBUG", "false").strip().lower() in ("1", "true", "yes")

# ──────────────────────────────────────────────
# ANSI colour codes for terminal readability
# ──────────────────────────────────────────────
_RESET  = "\033[0m"
_BOLD   = "\033[1m"
_DIM    = "\033[2m"
_GREEN  = "\033[92m"
_YELLOW = "\033[93m"
_RED    = "\033[91m"
_BLUE   = "\033[94m"
_CYAN   = "\033[96m"
_MAGENTA= "\033[95m"

# ──────────────────────────────────────────────
# Custom formatter
# ──────────────────────────────────────────────
class _ColourFormatter(logging.Formatter):
    LEVEL_COLOURS = {
        "DEBUG":    _CYAN,
        "INFO":     _BLUE,
        "WARNING":  _YELLOW,
        "ERROR":    _RED,
        "CRITICAL": _MAGENTA,
    }

    def format(self, record: logging.LogRecord) -> str:
        ts    = datetime.now().strftime("%H:%M:%S")
        level = record.levelname
        col   = self.LEVEL_COLOURS.get(level, "")
        tag   = f"{_DIM}[{ts}]{_RESET} {col}{_BOLD}[{level[:4]}]{_RESET}"
        # Prefix tag comes from record.name (e.g. "BG", "COSMOS", "LEAVE")
        prefix = f" {_DIM}{record.name}{_RESET}" if record.name != "root" else ""
        return f"{tag}{prefix} {record.getMessage()}"


def _build_logger() -> logging.Logger:
    logger = logging.getLogger("hamdaz")
    if logger.handlers:           # avoid duplicate handlers on reload
        return logger
    logger.setLevel(logging.DEBUG if DEBUG_MODE else logging.INFO)

    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(_ColourFormatter())
    handler.setLevel(logging.DEBUG)
    logger.addHandler(handler)
    logger.propagate = False
    return logger


_root_logger = _build_logger()


# ──────────────────────────────────────────────
# Public helper — thin wrapper with named tags
# ──────────────────────────────────────────────
class _TaggedLogger:
    """
    Provides log.info / log.debug / log.warn / log.error / log.success
    and supports an optional [tag] prefix:

        log.info("Started", tag="BG")
        log.error("Failed", tag="COSMOS", exc=e)
    """

    def _emit(self, level: str, msg: str, tag: str = "", exc: Exception = None) -> None:
        child = _root_logger.getChild(tag) if tag else _root_logger
        child.setLevel(logging.DEBUG)
        full_msg = str(msg)
        if exc:
            full_msg += f" | {type(exc).__name__}: {exc}"
        getattr(child, level)(full_msg)

    # Always shown
    def info(self, msg: str, tag: str = "", exc: Exception = None) -> None:
        self._emit("info", msg, tag, exc)

    def warn(self, msg: str, tag: str = "", exc: Exception = None) -> None:
        self._emit("warning", msg, tag, exc)

    def error(self, msg: str, tag: str = "", exc: Exception = None) -> None:
        self._emit("error", msg, tag, exc)

    # Only shown when DEBUG=true
    def debug(self, msg: str, tag: str = "", exc: Exception = None) -> None:
        if DEBUG_MODE:
            self._emit("debug", msg, tag, exc)

    def success(self, msg: str, tag: str = "", exc: Exception = None) -> None:
        """Positive confirmation — shown only in debug mode to reduce noise."""
        if DEBUG_MODE:
            _root_logger.getChild(tag or "root").info(f"{_GREEN}✔{_RESET} {msg}")


log = _TaggedLogger()

# ──────────────────────────────────────────────
# Announce mode on import
# ──────────────────────────────────────────────
if DEBUG_MODE:
    log.debug("Debug logging ENABLED — set DEBUG=false in .env to silence verbose output", tag="LOGGER")
else:
    log.info("Logging ready (verbose off — set DEBUG=true in .env to enable debug output)", tag="LOGGER")
