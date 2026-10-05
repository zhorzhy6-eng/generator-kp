# -*- coding: utf-8 -*-
"""Временная проверка: собираем настоящий отчёт «📋 Диагностика».

Tkinter подменяется заглушкой: GUI здесь не нужен, а код отчёта — тот же,
что выполняется в программе.
"""
import sys
import types
from pathlib import Path

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "src"))

# ---------- Заглушка tkinter ----------
class _PermissiveModule(types.ModuleType):
    """Модуль-заглушка: любой запрошенный виджет создаётся на лету."""

    def __getattr__(self, name):
        if name.startswith("__"):
            raise AttributeError(name)

        class _Any(_Widget):
            pass

        _Any.__name__ = name
        setattr(self, name, _Any)
        return _Any


class _Widget:
    def __init__(self, *a, **k):
        self._packed = False

    def pack(self, *a, **k):
        self._packed = True

    def pack_forget(self):
        self._packed = False

    def winfo_manager(self):
        return "pack" if self._packed else ""

    def get(self, *a, **k):
        return "0.8"

    def set(self, *a, **k):
        pass

    def bind(self, *a, **k):
        pass

    def config(self, *a, **k):
        pass

    def configure(self, *a, **k):
        pass

    def after(self, delay, callback=None, *a):
        return "after-id"

    def __getattr__(self, name):
        return lambda *a, **k: None


tk = _PermissiveModule("tkinter")
tk.constants = _PermissiveModule("tkinter.constants")
for const in ("END", "LEFT", "RIGHT", "TOP", "BOTTOM", "BOTH", "X", "Y", "W", "E", "N", "S", "NW", "NE", "SW", "SE", "CENTER", "VERTICAL", "HORIZONTAL", "DISABLED", "NORMAL", "WORD", "INSERT", "ANCHOR", "SUNKEN", "RAISED", "GROOVE", "RIDGE", "FLAT", "SOLID", "ACTIVE"):
    setattr(tk.constants, const, const)
sys.modules["tkinter.constants"] = tk.constants
for sub in ("ttk", "messagebox", "filedialog", "scrolledtext", "font"):
    fake = _PermissiveModule(f"tkinter.{sub}")
    sys.modules[f"tkinter.{sub}"] = fake
    setattr(tk, sub, fake)  # чтобы tk.ttk.Notebook тоже работал
sys.modules["tkinter"] = tk

import time  # noqa: E402

_t0 = time.time()
print("[тайминг] до импорта GUI: %.2f s" % (time.time() - _t0))
import генератор_кп_gigachat as gui  # noqa: E402

print("[тайминг] модуль импортирован: %.2f s" % (time.time() - _t0))
_t1 = time.time()
status = gui.get_status_info()
print("[тайминг] get_status_info(): %.2f s" % (time.time() - _t1))
_t1 = time.time()
gui.is_gigachat_configured()
print("[тайминг] is_gigachat_configured(): %.2f s" % (time.time() - _t1))

print("=" * 72)
print("ОТЧЁТ «📋 ДИАГНОСТИКА» (как его увидит пользователь)")
print("=" * 72)
report = gui._build_diagnostics_report()
print(report)

print()
print("=" * 72)
print("ТЕКСТ ОКНА «Как исправить» (кнопка на SSL-плашке)")
print("=" * 72)
print(gui.get_ssl_help())

print()
print("=" * 72)
print("СОСТОЯНИЕ ПЛАШЕК")
print("=" * 72)
print("  сертификат найден (_certificate_present):", gui._certificate_present())
print("  ключ задан (GIGACHAT_READY):", gui.GIGACHAT_READY)
print("  SSL-ошибка была (has_ssl_error):", gui.has_ssl_error())
print("  плашка «Ключ не задан» упакована:", bool(gui.key_warning_label.winfo_manager()))
