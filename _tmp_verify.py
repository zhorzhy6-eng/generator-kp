# -*- coding: utf-8 -*-
"""
Временная проверка Шагов 4, 5 и 8 (не входит в проект, удаляется после).

Что делает:
  1. Компилирует модули проекта (.pyc пишется в TEMP — папка проекта
     может быть недоступна на запись в этой среде).
  2. Проверяет generate_gigachat(timeout=...) — команда из ТЗ.
  3. Пишет в лог полный контекст SSL-ошибки.
  4. Импортирует GUI-модуль, проверяет плашку SSL при старте и
     собирает отчёт диагностики.
"""

import os
import sys
import tempfile
from pathlib import Path

PROJECT = Path(r"E:\Prpgrammy\Шаблоны")
SRC = PROJECT / "src"
TMP = Path(tempfile.gettempdir()) / "kp_verify"

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
sys.dont_write_bytecode = True

print("=" * 70)
print("1. КОМПИЛЯЦИЯ МОДУЛЕЙ")
print("=" * 70)

import py_compile

TMP.mkdir(parents=True, exist_ok=True)
for name in (
    "генератор_кп_gigachat.py",
    "llm_provider.py",
    "logger_config.py",
    "lot_parser.py",
):
    path = SRC / name
    if not path.exists():
        print(f"   ⏭️  {name}: файла нет — пропускаю")
        continue
    cfile = TMP / (name + ".pyc")
    try:
        py_compile.compile(str(path), cfile=str(cfile), doraise=True)
        print(f"   ✅ {name}: синтаксис OK")
    except py_compile.PyCompileError as exc:
        print(f"   ❌ {name}: {exc}")
        sys.exit(1)

print()
print("=" * 70)
print("2. GENERATE_GIGACHAT С TIMEOUT (команда из ТЗ)")
print("=" * 70)

sys.path.insert(0, str(SRC))

import llm_provider as lp

result = lp.generate_gigachat("тест", timeout=10)
print(f"   result: {type(result)} {repr(result)[:100]}")
print(f"   исключения не вылетели: ✅")
print(f"   SSL-ошибка в сессии: {'⚠️ да' if lp.has_ssl_error() else 'нет'}")

status = lp.get_ca_bundle_status()
print(f"   сертификат найден: {status['found'] or 'НЕТ'}")
print(f"   ожидаемый путь: {status['expected']}")
print(f"   verify_ssl: {status['verify_ssl']}")

print()
print("=" * 70)
print("3. ЛОГ: КОНТЕКСТ SSL-ОШИБКИ")
print("=" * 70)

import logger_config

log_file = Path(logger_config.LOG_FILE)
lines = log_file.read_text(encoding="utf-8", errors="replace").splitlines()
ssl_lines = [ln for ln in lines if "SSL" in ln or "Traceback" in ln]
print(f"   лог: {log_file}")
print(f"   строк с SSL/traceback: {len(ssl_lines)}")
for ln in ssl_lines[-14:]:
    print("   |", ln[:150])

print()
print("=" * 70)
print("4. GUI: ПЛАШКА SSL И ДИАГНОСТИКА")
print("=" * 70)

import tkinter as tk

import генератор_кп_gigachat as gui  # noqa: E402  (импорт после проверок)

print(f"   окно создано: {gui.window.winfo_exists() == 1}")
print(f"   GIGACHAT_READY: {gui.GIGACHAT_READY}")
print(f"   сертификат есть: {gui._certificate_present()}")

gui.window.update_idletasks()
banner_shown = bool(gui.ssl_banner.winfo_manager())
expected_banner = gui.GIGACHAT_READY and not gui._certificate_present()
print(f"   плашка SSL показана: {banner_shown} (ожидалось: {expected_banner})")
print(f"   текст плашки: {gui.ssl_label.cget('text')[:80]!r}")

print()
print("   --- ОТЧЁТ ДИАГНОСТИКИ ---")
report = gui._build_diagnostics_report()
for ln in report.splitlines():
    print("   |", ln)

checks = {
    ".env в отчёте": ".env найден" in report,
    "ключ с маской": "символов, значение скрыто" in report,
    "SSL-раздел с URL": "russian_trusted_root_ca.cer" in report,
    "пути поиска сертификата": "Где искали сертификат" in report,
    "логгер": "Логгер" in report,
    "мьютекс": "Мьютекс" in report,
    "python": "Python:" in report,
    "tkinter": "tkinter: OK" in report,
    "python-docx": "python-docx: OK" in report,
    "pyperclip": "pyperclip: OK" in report,
    "итог": "Итог" in report,
}

print()
print("   --- САМОПРОВЕРКА ОТЧЁТА ---")
all_ok = True
for name, ok in checks.items():
    print(f"   {'✅' if ok else '❌'} {name}")
    all_ok = all_ok and ok

print()
print("=" * 70)
print("4b. ПЛАШКА ПРИ СТАРТЕ (без предшествующих ошибок)")
print("=" * 70)

import subprocess

code = (
    "import sys, os; sys.path.insert(0, r'%s');"
    "sys.dont_write_bytecode = True;"
    "import генератор_кп_gigachat as g;"
    "print('@@BANNER_SHOWN=', bool(g.ssl_banner.winfo_manager()));"
    "print('@@BANNER_TEXT=', str(g.ssl_label.cget('text')));"
    "g.window.destroy()" % SRC
)
env = dict(
    os.environ, PYTHONDONTWRITEBYTECODE="1", PYTHONIOENCODING="utf-8", PYTHONUTF8="1"
)
proc = subprocess.run(
    [sys.executable, "-c", code],
    cwd=str(PROJECT),
    env=env,
    capture_output=True,
    text=True,
    encoding="utf-8",
    errors="replace",
)
for line in (proc.stdout or "").splitlines():
    if line.startswith("@@BANNER"):
        print("   ", line.replace("@@", ""))
print("   старт без ошибок:", proc.returncode == 0)
if proc.returncode != 0:
    print("   --- stderr ---")
    for line in (proc.stderr or "").splitlines()[-25:]:
        print("   |", line)

print()
print("=" * 70)
print("5. ДИАЛОГИ (создаются и закрываются без ошибок)")
print("=" * 70)

try:
    gui.show_ssl_help()
    gui.window.update_idletasks()
    print("   ✅ show_ssl_help(): окно создано")
    for child in gui.window.winfo_children():
        if isinstance(child, tk.Toplevel) or child.winfo_class() == "Toplevel":
            child.destroy()
except Exception as exc:
    print(f"   ❌ show_ssl_help(): {exc}")
    all_ok = False

try:
    gui.show_diagnostics()
    gui.window.update_idletasks()
    print("   ✅ show_diagnostics(): окно создано, отчёт собирается в фоне")
except Exception as exc:
    print(f"   ❌ show_diagnostics(): {exc}")
    all_ok = False

print()
print("=" * 70)
print(f"ИТОГ ПРОВЕРКИ: {'✅ ВСЁ ОК' if all_ok else '❌ ЕСТЬ ПРОБЛЕМЫ'}")
print("=" * 70)

try:
    gui.window.destroy()
except Exception:
    pass

sys.exit(0 if all_ok else 1)
