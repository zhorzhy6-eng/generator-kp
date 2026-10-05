# -*- coding: utf-8 -*-
"""Временная диагностика №3: кто закрывает окно.

Перехватываем destroy/quit/mainloop и пишем стек вызовов того, кто
инициировал закрытие окна.
"""
import os
import sys
import tempfile
import threading
import time
import traceback

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src")
TMP_LOGS = os.path.join(tempfile.gettempdir(), "kp_probe_logs")
os.makedirs(TMP_LOGS, exist_ok=True)
OUT = os.path.join(TMP_LOGS, "probe_result2.txt")
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
logger_config.LOG_FILE = os.path.join(TMP_LOGS, "probe_generator2.log")

import tkinter as tk  # noqa: E402

real_Tk = tk.Tk
real_Misc_destroy = tk.Misc.destroy
real_Misc_quit = tk.Misc.quit

state = {"mainloop_entered": False, "close_reason": None}


class SpyTk(real_Tk):
    def destroy(self):
        if state["mainloop_entered"] and state["close_reason"] is None:
            state["close_reason"] = "Tk.destroy()\n" + "".join(traceback.format_stack())
            note("!!! window.destroy() called from:\n" + "".join(traceback.format_stack()))
        return super().destroy()

    def quit(self):
        if state["mainloop_entered"] and state["close_reason"] is None:
            state["close_reason"] = "Tk.quit()\n" + "".join(traceback.format_stack())
            note("!!! window.quit() called from:\n" + "".join(traceback.format_stack()))
        return super().quit()

    def mainloop(self, *a, **k):
        note("mainloop ENTER")
        state["mainloop_entered"] = True
        try:
            return super().mainloop(*a, **k)
        finally:
            state["mainloop_entered"] = False
            note("mainloop EXIT")


def spy_destroy(self):
    if state["mainloop_entered"] and state["close_reason"] is None:
        state["close_reason"] = "Misc.destroy on %r\n" % (self,) + "".join(
            traceback.format_stack()
        )
        note("!!! Misc.destroy on %r from:\n%s" % (self, "".join(traceback.format_stack())))
    return real_Misc_destroy(self)


def spy_quit(self):
    if state["mainloop_entered"] and state["close_reason"] is None:
        state["close_reason"] = "Misc.quit on %r\n" % (self,) + "".join(
            traceback.format_stack()
        )
        note("!!! Misc.quit on %r from:\n%s" % (self, "".join(traceback.format_stack())))
    return real_Misc_quit(self)


tk.Tk = SpyTk
tk.Misc.destroy = spy_destroy
tk.Misc.quit = spy_quit

note("importing generator...")
import генератор_кп_gigachat as g  # noqa: E402

note("imported; window=%r" % (g.window,))


def watchdog():
    time.sleep(20)
    note("watchdog: mainloop still running after 20s -> OK?! reason=%s" % state["close_reason"])
    try:
        note("watchdog: exists=%s viewable=%s" % (
            g.window.winfo_exists(), g.window.winfo_viewable()))
    except Exception as exc:
        note("watchdog: state read failed: %r" % (exc,))
    os._exit(0)


threading.Thread(target=watchdog, daemon=True).start()

note("calling g.window.mainloop()")
g.window.mainloop()
note("g.window.mainloop() RETURNED. reason=%s" % state["close_reason"])
os._exit(9)
