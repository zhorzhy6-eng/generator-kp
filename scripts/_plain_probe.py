# -*- coding: utf-8 -*-
"""Проба A: обычный импорт генератора БЕЗ подмены tk.Tk."""
import os
import sys
import tempfile

OUT = os.path.join(tempfile.gettempdir(), "kp_probe_logs", "plain.txt")
os.makedirs(os.path.dirname(OUT), exist_ok=True)
lines = []


def note(m):
    lines.append(str(m))
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))
os.chdir(ROOT)

import logger_config  # noqa: E402

logger_config.LOG_DIR = os.path.join(tempfile.gettempdir(), "kp_probe_logs")
logger_config.LOG_FILE = os.path.join(logger_config.LOG_DIR, "plain.log")

note("importing generator WITHOUT SpyTk")
import генератор_кп_gigachat as g  # noqa: E402

note("import RETURNED; mainloop called = %s" % getattr(g, "_MAINLOOP_CALLED", "no-attr"))
note("window exists = %s" % g.window.winfo_exists())
note("protocol = %r" % g.window.protocol("WM_DELETE_WINDOW"))
note("PLAIN CHECK DONE")
os._exit(0)
