# -*- coding: utf-8 -*-
"""Временная диагностика №4: проверка после правок.

Импортирует обновлённый модуль генератора, проверяет что окно создаётся,
глобальный excepthook установлен, а mainloop работает.
Результат пишется файлом (stdout под песочницей может блокироваться).
"""
import os
import sys
import tempfile
import threading
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src")
TMP_LOGS = os.path.join(tempfile.gettempdir(), "kp_probe_logs")
os.makedirs(TMP_LOGS, exist_ok=True)
OUT = os.path.join(TMP_LOGS, "probe_result3.txt")
sys.path.insert(0, SRC)
os.chdir(ROOT)

steps = []
T0 = time.time()


def note(msg):
    steps.append("%.2f  %s" % (time.time() - T0, msg))
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(steps) + "\n")


import logger_config  # noqa: E402

logger_config.LOG_DIR = TMP_LOGS
logger_config.LOG_FILE = os.path.join(TMP_LOGS, "probe_generator3.log")

import tkinter as tk  # noqa: E402

# Перехватываем создание окна, чтобы не запускать mainloop модуля
real_Tk = tk.Tk
created = []


class SpyTk(real_Tk):
    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        created.append(self)


tk.Tk = SpyTk

note("importing generator...")
import генератор_кп_gigachat as g  # noqa: E402

note("imported OK; GIGACHAT_READY=%s" % g.GIGACHAT_READY)
note("excepthook installed = %s" % (sys.excepthook is g._excepthook))
note("threading.excepthook set = %s" % (threading.excepthook is g._thread_excepthook))
note("window protocol WM_DELETE_WINDOW = %r" % g.window.protocol("WM_DELETE_WINDOW"))
note("has update_ssl_banner = %s" % callable(g.update_ssl_banner))
note("has show_ssl_help = %s" % callable(g.show_ssl_help))
note("has show_diagnostics = %s" % callable(g.show_diagnostics))
note("documents dir = %s" % g._documents_dir())

# Проверяем плашку SSL (должна быть скрыта, пока ошибки не было)
g.update_ssl_banner()
note("ssl banner mapped (expect False) = %s" % g.ssl_banner.winfo_ismapped())

# Симулируем SSL-ошибку
import llm_provider  # noqa: E402

llm_provider.note_ssl_error("ConnectError: [SSL: CERTIFICATE_VERIFY_FAILED] test")
g.update_ssl_banner()
note("ssl banner mapped after error (expect True) = %s" % g.ssl_banner.winfo_ismapped())

# Проверяем, что обработка ключа идёт в правильный путь
note("env write path = %s" % llm_provider.get_env_write_path())

# Проверяем mainloop на скрытом окне
hidden = created[0]
hidden.withdraw()


def stop():
    time.sleep(2)
    note("mainloop ran fine for 2s -> destroying")
    hidden.destroy()


threading.Thread(target=stop, daemon=True).start()
t0 = time.time()
hidden.mainloop()
note("mainloop blocked for %.2fs (expect ~2s)" % (time.time() - t0))

# Проверяем on_closing
g.on_closing()
note("on_closing() executed OK; closed flag = %s" % g._window_closed["value"])

note("ALL CHECKS DONE")
os._exit(0)
