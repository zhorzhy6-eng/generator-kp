# -*- coding: utf-8 -*-
"""Минимальные пробы отдельных вызовов. Результат — файл."""
import os
import sys
import tempfile

OUT = os.path.join(tempfile.gettempdir(), "kp_probe_logs", "mini.txt")
os.makedirs(os.path.dirname(OUT), exist_ok=True)
lines = []


def note(m):
    lines.append(str(m))
    with open(OUT, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")


note("start")
import tkinter as tk

note("tkinter imported")
root = tk.Tk()
note("Tk created")
root.withdraw()


def cb():
    pass


note("step: protocol()")
root.protocol("WM_DELETE_WINDOW", cb)
note("protocol() OK -> %r" % root.protocol("WM_DELETE_WINDOW"))

note("step: after()")
root.after(100, cb)
note("after() OK")

note("step: ctypes mutex")
try:
    import ctypes
    from ctypes import wintypes

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    k32.CreateMutexW.restype = wintypes.HANDLE
    note("WinDLL attributes set")
    h = k32.CreateMutexW(None, True, "Global\\GeneratorKP_ProbeMini")
    note("CreateMutexW -> %r last_error=%s" % (h, ctypes.get_last_error()))
except Exception as exc:
    import traceback

    note("ctypes FAILED: %r\n%s" % (exc, traceback.format_exc()))

note("step: update()")
root.update()
note("update() OK")

note("ALL MINI CHECKS DONE")
sys.stdout.flush()
os._exit(0)
