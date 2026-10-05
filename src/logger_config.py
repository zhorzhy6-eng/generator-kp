#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Единая настройка логирования для «Генератора КП».

Задачи модуля:
  * Все пользовательские данные (settings.json, логи) лежат в КОРНЕ ПРОЕКТА,
    а не в текущей рабочей директории, — поэтому программа работает одинаково
    при любом способе запуска: из меню .bat, напрямую из src/ или из .exe.
  * Логирование НИКОГДА не роняет программу. Если папку проекта нельзя
    писать (запуск из Program Files, второй экземпляр, права, антивирус) —
    логи уезжают в %USERPROFILE%\\GeneratorKP\\logs, а программа продолжает
    работать. Раньше в этой ситуации был мгновенный вылет ещё до появления
    окна, причём без единого сообщения пользователю.
  * Файл лога открывается ЛЕНИВО (delay=True): при импорте модуля файл не
    захватывается. Это важно, потому что два одновременно запущенных
    экземпляра программы иначе блокируют друг друга.
"""

import logging
import os
import sys
import tempfile
from logging.handlers import RotatingFileHandler
from datetime import datetime

# ============================================
# БАЗОВЫЕ ПУТИ ПРОЕКТА
# ============================================


def _project_root() -> str:
    """Возвращает корень проекта независимо от текущей рабочей директории."""
    if getattr(sys, "frozen", False):
        # Собранный PyInstaller-файл: пользовательские данные рядом с .exe
        return os.path.dirname(os.path.abspath(sys.executable))
    # Исходники: этот файл лежит в <корень проекта>/src/
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


PROJECT_ROOT = _project_root()
SETTINGS_PATH = os.path.join(PROJECT_ROOT, "settings.json")
LOG_DIR = os.path.join(PROJECT_ROOT, "logs")
LEGACY_LOG_FILE = os.path.join(PROJECT_ROOT, "generator_kp.log")

# ============================================
# ПОИСК ПАПКИ ДЛЯ ЛОГОВ, КОТОРУЮ МОЖНО ПИСАТЬ
# ============================================


def _user_log_dir() -> str:
    """
    Запасная папка для логов в профиле пользователя.

    Нужна, когда папка проекта недоступна на запись: .exe лежит в
    Program Files, папка помечена «только чтение», или её держит антивирус.
    """
    base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    return os.path.join(base, "GeneratorKP", "logs")


def _last_resort_log_dir() -> str:
    """Совсем запасной вариант — временная папка пользователя."""
    return os.path.join(tempfile.gettempdir(), "GeneratorKP", "logs")


def _ensure_writable_dir(path: str) -> bool:
    """
    Проверяет, что папку можно создать и в неё можно писать.

    Проверка настоящая (создаём и удаляем пробный файл), потому что одних
    прав на папку мало: файл может держать другой процесс или антивирус.

    ВАЖНО: пробный файл уникален для процесса — иначе два экземпляра
    программы, стартовавшие одновременно, мешали бы друг другу.
    """
    try:
        os.makedirs(path, exist_ok=True)
        probe = os.path.join(path, f".write_test_{os.getpid()}")
        with open(probe, "w", encoding="utf-8") as fh:
            fh.write("ok")
        os.remove(probe)
        return True
    except Exception:
        return False


def _pick_log_dir() -> tuple:
    """
    Выбирает первую доступную на запись папку для логов.

    Returns:
        (путь к папке, причина выбора) — причина попадает в лог, чтобы
        было понятно, почему логи лежат не там, где ожидалось.
    """
    candidates = (
        (LOG_DIR, "папка проекта"),
        (_user_log_dir(), "папка пользователя (%LOCALAPPDATA%\\GeneratorKP)"),
        (_last_resort_log_dir(), "временная папка (%TEMP%)"),
    )
    for path, reason in candidates:
        if _ensure_writable_dir(path):
            return path, reason
    # Не должно случиться, но если совсем ничего не пишется — пишем в TEMP
    return tempfile.gettempdir(), "системная временная папка"


def _make_formatter() -> logging.Formatter:
    return logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def _make_file_handler(path: str):
    """
    Создаёт обработчик файла лога.

    delay=True — файл реально открывается при первой записи, а не при
    импорте модуля. Так запуск программы не захватывает файл заранее и
    не мешает другим экземплярам.
    """
    handler = RotatingFileHandler(
        path,
        maxBytes=10 * 1024 * 1024,
        backupCount=5,
        encoding="utf-8",
        delay=True,
    )
    handler.setFormatter(_make_formatter())
    return handler


def _init_paths() -> tuple:
    """
    Выбирает папку и файл лога при импорте модуля.

    Returns:
        (папка логов, файл лога, предупреждение или None)
    """
    if _ensure_writable_dir(LOG_DIR):
        return LOG_DIR, os.path.join(LOG_DIR, _log_name()), None

    fallback_dir, reason = _pick_log_dir()
    warning = (
        f"Папка проекта {LOG_DIR} недоступна на запись — логи перенесены в "
        f"{fallback_dir} ({reason})"
    )
    return fallback_dir, os.path.join(fallback_dir, _log_name()), warning


def _log_name() -> str:
    return f"generator_{datetime.now().strftime('%Y%m%d')}.log"


# Основной лог-файл выбирается один раз при импорте модуля
LOG_DIR, LOG_FILE, _LOG_WARNING = _init_paths()


def get_diagnostics() -> dict:
    """Сведения о логировании — для лога и для диагностики проблем запуска."""
    return {
        "project_root": PROJECT_ROOT,
        "log_dir": LOG_DIR,
        "log_file": LOG_FILE,
        "settings_path": SETTINGS_PATH,
        "fallback_warning": _LOG_WARNING,
        "frozen": bool(getattr(sys, "frozen", False)),
    }


# ============================================
# НАСТРОЙКА ЛОГГЕРА
# ============================================

_configured: dict = {}


def _make_console_handler() -> logging.Handler:
    """
    Консольный вывод в UTF-8.

    В .exe потока вывода нет (--noconsole), поэтому обработчик просто
    не добавляется. В консоли Windows переключаем кодировку на UTF-8,
    иначе русский текст и эмодзи падают с UnicodeEncodeError.
    """
    if sys.stderr is None:
        return None
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    try:
        return logging.StreamHandler()
    except Exception:
        return None


def setup_logger(name: str, level=logging.INFO) -> logging.Logger:
    """
    Настраивает логгер для модуля.

    Все логи пишутся в ОДИН файл с ротацией: каждый логгер получает свои
    обработчики, указывающие на тот же файл. Propagation при этом глушится
    для логгеров с собственными обработчиками — иначе каждая запись
    попадала бы в файл дважды.

    Обработчики создаются заново только если их ещё нет, поэтому повторные
    вызовы (например, setup_logger из llm_provider) безопасны.
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)

    if logger.handlers:
        # У этого логгера уже есть обработчики — не дублируем
        return logger

    file_handler = None
    try:
        file_handler = _make_file_handler(LOG_FILE)
    except Exception:
        file_handler = None

    if file_handler is None:
        # Файл недоступен — пробуем папку пользователя
        fallback_dir, fallback_reason = _pick_log_dir()
        try:
            fallback_path = os.path.join(fallback_dir, _log_name())
            file_handler = _make_file_handler(fallback_path)
            globals()["LOG_DIR"] = fallback_dir
            globals()["LOG_FILE"] = fallback_path
            _configured["fallback_reason"] = fallback_reason
        except Exception:
            file_handler = None

    if file_handler is not None:
        logger.addHandler(file_handler)

    console_handler = _make_console_handler()
    if console_handler is not None:
        console_handler.setFormatter(_make_formatter())
        logger.addHandler(console_handler)

    # Обработчики вешаем на сам логгер, поэтому вверх записи не отдаём:
    # у корневого логгера обработчиков нет, и без этого строки бы терялись.
    logger.propagate = False

    # Ошибки логирования не должны ронять программу
    logger.raiseExceptions = False

    if not _configured.get("ready"):
        _configured["ready"] = True
        logger.info("Логирование: %s", LOG_FILE)
        if _configured.get("fallback_reason"):
            logger.warning(
                "Папка проекта недоступна — логи перенесены в %s",
                _configured["fallback_reason"],
            )
        if file_handler is None:
            logger.error("Не удалось открыть файл лога — вывод только в консоль")

    return logger
