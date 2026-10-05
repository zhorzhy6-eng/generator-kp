# -*- coding: utf-8 -*-
"""Проба после правок: импорт больше не захватывает окно."""
import os
import sys
import tempfile
import threading
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TMP_LOGS = os.path.join(tempfile.gettempdir(), "kp_probe_logs")
os.makedirs(TMP_LOGS, exist_ok=True)
OUT = os.path.join(TMP_LOGS, "final.txt")
sys.path.insert(0, os.path.join(ROOT, "src"))
os.chdir(ROOT)

lines = []


def note(m):
    lines.append("%.2f  %s" % (time.time() - T0, m))
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")


T0 = time.time()
note("start")

import logger_config  # noqa: E402

logger_config.LOG_DIR = TMP_LOGS
logger_config.LOG_FILE = os.path.join(TMP_LOGS, "final.log")

note("importing generator (must NOT hang)")
import генератор_кп_gigachat as g  # noqa: E402

note("import RETURNED")
note("GIGACHAT_READY = %s" % g.GIGACHAT_READY)
note("excepthook installed = %s" % (sys.excepthook is g._excepthook))
note("threading excepthook = %s" % (threading.excepthook is g._thread_excepthook))
note("protocol = %r" % g.window.protocol("WM_DELETE_WINDOW"))
note("documents dir = %s" % g._documents_dir())
note("env write path = %s" % g.get_env_write_path())

note("ssl banner manager before (expect none) = %r" % g.ssl_banner.winfo_manager())
import llm_provider  # noqa: E402

llm_provider.note_ssl_error("ConnectError: [SSL: CERTIFICATE_VERIFY_FAILED] test")
g.update_ssl_banner()
note("ssl banner manager after SSL error (expect pack) = %r" % g.ssl_banner.winfo_manager())
llm_provider.clear_ssl_error()
g.update_ssl_banner()
note("ssl banner manager cleared (expect none) = %r" % g.ssl_banner.winfo_manager())

h, already = g._acquire_single_instance_lock()
note("mutex: handle=%r already_running=%s" % (bool(h), already))
if h:
    import ctypes

    ctypes.WinDLL("kernel32").CloseHandle(h)
    note("mutex released")

note("on_closing()")
g.on_closing()
note("window destroyed flag = %s" % g._window_closed["value"])

note("ALL CHECKS PASSED")
os._exit(0)
