"""
colprint.py — tiny coloured-print helpers for pipeline debugging.

Usage:
    import colprint as cp
    cp.info("Reading files...")        # cyan
    cp.ok("12 samples recovered")      # green
    cp.warn("sample is partial")       # yellow
    cp.err("Prism timed out")          # red
    cp.section("STEP 3: Prism")        # bold

Colours are automatic: plain text when the output is not a terminal or when
the NO_COLOR environment variable is set, so redirecting to a log file
stays clean. No dependencies — ANSI codes work in Terminal.app, iTerm2,
and most modern terminals.

If you ever want to force colours off in one run:
    NO_COLOR=1 python3 my_script.py
"""

import os
import sys

_RESET = "\x1b[0m"
_CODES = {
    "info": "\x1b[36m",     # cyan
    "ok": "\x1b[32m",       # green
    "warn": "\x1b[33m",     # yellow
    "err": "\x1b[31m",      # red
    "bold": "\x1b[1m",      # bold (keeps current colour)
    "blue": "\x1b[34m",
    "magenta": "\x1b[35m",
}


def _enabled():
    if os.environ.get("NO_COLOR"):
        return False
    try:
        return sys.stdout.isatty()
    except Exception:
        return False


_ENABLED = _enabled()


def _p(color, msg):
    if _ENABLED:
        print(f"{_CODES[color]}{msg}{_RESET}")
    else:
        print(msg)


def info(msg):
    _p("info", msg)


def ok(msg):
    _p("ok", msg)


def warn(msg):
    _p("warn", msg)


def err(msg):
    _p("err", msg)


def section(msg):
    _p("bold", f"== {msg} ==")
