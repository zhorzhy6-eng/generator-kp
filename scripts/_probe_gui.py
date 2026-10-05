# -*- coding: utf-8 -*-
"""Временная диагностика №2.

Разделяет два вопроса:
  A) импорт модуля + построение GUI   -> доходит ли до mainloop
  B) реально ли mainloop() удерживает управление

Результат пишется ФАЙЛОМ (stdout под песочницей может блокироваться).
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
OUT = os.path.join(TMP_LOGS, "probe_result.txt")
sys.path.insert(0, SRC)
os.chdir(ROOT)

steps = []


def note(msg):
    steps.append("%.2f  %s" % (time.time() - T0, msg))
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(steps) + "\n")


T0 = time.time()
note("start; tmp=%s" % TMP_LOGS)

import logger_config  # noqa: E402

logger_config.LOG_DIR = TMP_LOGS
logger_config.LOG_FILE = os.path.join(TMP_LOGS, "probe_generator.log")
note("logger_config patched")

import tkinter as tk  # noqa: E402

note("tkinter imported")

kw = {"calls": []}
real_Tk = tk.Tk


class SpyTk(real_Tk):
    def __init__(self, *a, **k):
        kw["calls"].append("Tk()")
        note("Tk() called")
        super().__init__(*a, **k)

    def mainloop(self, *a, **k):
        note("mainloop() ENTER")
        try:
            r = super().mainloop(*a, **k)
        except BaseException as exc:
            note("mainloop() RAISED %r" % (exc,))
            raise
        note("mainloop() RETURNED")
        return r


tk.Tk = SpyTk

note("importing generator module...")
import генератор_кп_gigachat as g  # noqa: E402

note("module imported; GIGACHAT_READY=%s" % g.GIGACHAT_READY)
notify_calls = [c for c in kw["calls"]]
note("Tk instantiations: %d" % len(notify_calls))

# A) Скрытый второй root: проверяем, что mainloop вообще блокирует
hidden = tk.Toplevel()
hidden.withdraw()
note("about to test hidden mainloop")


def quit_hidden():
    note("hidden after() fired")
    hidden.quit()


hidden.after(1500, quit_hidden)
t_hidden = time.time()
hidden.mainloop()
note("hidden mainloop blocked for %.2fs (ok if ~1.5)" % (time.time() - t_hidden))


# B) Проверяем окно генератора: живо ли оно
result = {}


def watchdog():
    time.sleep(5)
    try:
        result["exists"] = g.window.winfo_exists()
        result["viewable"] = g.window.winfo_viewable()
        result["geom"] = g.window.winfo_geometry()
        result["mapped"] = g.window.winfo_ismapped()
    except Exception as exc:
        result["error"] = repr(exc)
    note("WINDOW STATE: %s" % result)
    os._exit(0)


threading.Thread(target=watchdog, daemon=True).start()
note("entering generator window.mainloop()")
g.window.mainloop()
note("!!! generator mainloop RETURNED immediately — окно закрылось само")
os._exit(7)
