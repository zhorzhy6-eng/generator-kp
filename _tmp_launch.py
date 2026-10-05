# -*- coding: utf-8 -*-
"""
Временная проверка Шага 8: реальный запуск GUI-модуля и корректное закрытие.

Окно реально создаётся и отрисовывается, поэтому проверяются:
  * старт без исключений;
  * вкладки Notebook (первая — «Юга»);
  * плашка SSL сразу при старте;
  * актуальный LOG_FILE из logger_config.
Закрытие — через on_closing(), то есть штатным путём (destroy + запись в лог).

Не входит в проект: удаляется после проверки.
"""

import sys
from pathlib import Path

PROJECT = Path(r"E:\Prpgrammy\Шаблоны")
sys.path.insert(0, str(PROJECT / "src"))
sys.dont_write_bytecode = True

import logger_config  # noqa: E402

import генератор_кп_gigachat as gui  # noqa: E402

gui.window.update()
gui.window.update_idletasks()

tabs = []
for tab_id in gui.notebook.tabs():
    tabs.append(gui.notebook.tab(tab_id, "text"))

print("RESULT окно создано:", gui.window.winfo_exists() == 1)
print("RESULT окно видимо:", gui.window.winfo_viewable() == 1)
print("RESULT вкладки:", tabs)
print("RESULT плашка SSL показана:", bool(gui.ssl_banner.winfo_manager()))
print("RESULT текст плашки:", str(gui.ssl_label.cget("text")))
print("RESULT LOG_FILE:", logger_config.LOG_FILE)

# Штатное закрытие — тот же путь, что по крестику и по кнопке «ВЫХОД»
gui.on_closing()

print("RESULT закрытие выполнено, окно существует:", gui.window.winfo_exists() == 1)
