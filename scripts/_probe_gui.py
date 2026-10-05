# -*- coding: utf-8 -*-
"""Временный диагностический скрипт. Удаляется после отладки.

Запускает главное окно генератора, ждёт 8 секунд и проверяет:
жив ли ещё mainloop. Весь stdout/stderr уходит в файл.
"""
import sys
import os
import time
import threading

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src"))
os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

print("PROBE: import start", flush=True)
import tkinter as tk  # noqa: E402

print("PROBE: tkinter OK", flush=True)

import генератор_кп_gigachat as g  # noqa: E402

print("PROBE: module imported OK", flush=True)
print("PROBE: GIGACHAT_READY =", g.GIGACHAT_READY, flush=True)

alive = {"value": None}


def check():
    t0 = time.time()
    print("PROBE: waiting for mainloop...", flush=True)
    while time.time() - t0 < 8:
        try:
            info = g.window.winfo_exists()
        except Exception as exc:
            print("PROBE: winfo_exists FAILED:", repr(exc), flush=True)
            break
        if not info:
            print("PROBE: window destroyed after %.2f s" % (time.time() - t0), flush=True)
            break
        time.sleep(0.5)
    else:
        print("PROBE: window ALIVE after 8 s -> mainloop is running", flush=True)
        alive["value"] = True

    try:
        print("PROBE: viewable =", g.window.winfo_viewable(), flush=True)
        print("PROBE: geometry =", g.window.winfo_geometry(), flush=True)
    except Exception as exc:
        print("PROBE: geometry read failed:", repr(exc), flush=True)

    print("PROBE: closing window now", flush=True)
    try:
        g.window.destroy()
    except Exception:
        pass


threading.Thread(target=check, daemon=True).start()

print("PROBE: entering mainloop", flush=True)
try:
    g.window.mainloop()
    print("PROBE: mainloop RETURNED normally", flush=True)
except SystemExit as exc:
    print("PROBE: SystemExit in mainloop:", repr(exc), flush=True)
except BaseException as exc:
    import traceback

    print("PROBE: EXCEPTION in mainloop:", repr(exc), flush=True)
    traceback.print_exc()
print("PROBE: script end", flush=True)
