import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from datetime import datetime

# ============================================
# БАЗОВЫЕ ПУТИ ПРОЕКТА
# ============================================
# Все пользовательские данные (settings.json, логи) лежат в корне проекта,
# а не в текущей рабочей директории, поэтому программы работают одинаково
# при любом способе запуска: из меню .bat, напрямую из src/ или из собранного .exe.


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

# Создаём папку для логов
os.makedirs(LOG_DIR, exist_ok=True)

# Путь к лог-файлу с датой
LOG_FILE = os.path.join(LOG_DIR, f"generator_{datetime.now().strftime('%Y%m%d')}.log")


def setup_logger(name: str, level=logging.INFO) -> logging.Logger:
    """
    Настраивает логгер для модуля.
    Все логи пишутся в один файл с ротацией.
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)

    # Очищаем старые обработчики, чтобы не дублировать
    if logger.handlers:
        logger.handlers.clear()

    # Формат лога
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Файловый обработчик с ротацией (макс 10 МБ, 5 бэкапов)
    file_handler = RotatingFileHandler(
        LOG_FILE, maxBytes=10*1024*1024, backupCount=5, encoding='utf-8'
    )
    file_handler.setFormatter(formatter)

    # Консольный обработчик
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger
