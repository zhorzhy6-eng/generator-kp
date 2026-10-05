#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Универсальный провайдер LLM для «Генератора КП».

Поддерживает два бэкенда:
  * GigaChat (облачный, Сбер) — ключ берётся ТОЛЬКО из .env
  * Ollama   (локальный)      — http://localhost:11434

Принципы безопасности:
  * API-ключ никогда не логируется: любые упоминания маскируются как "***"
  * Ключ не хранится в коде — только в .env (который в .gitignore)
  * SSL-верификация включена по умолчанию
  * Ни одна функция не выбрасывает исключение наружу: при ошибке возвращается None

Принципы скорости:
  * Клиент GigaChat создаётся лениво и кэшируется (singleton) — токен и
    TLS-соединение переиспользуются между запросами
  * Генерация идёт потоково (stream=True): текст начинает приходить сразу
  * Есть отдельная функция warmup() для прогрева авторизации в фоне
"""

import os
import sys
import re
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

# ============================================
# ЗАГРУЗКА .env (безопасно, если dotenv не установлен)
# ============================================
# Файл .env ищется в нескольких местах, потому что программа запускается
# по-разному: из исходников, из .bat и из собранного .exe. Внутри .exe
# переменная __file__ указывает во временную папку распаковки PyInstaller —
# искать .env там бессмысленно, поэтому рядом с .exe (sys.executable) файл
# проверяется отдельно и раньше всего.

IS_FROZEN = bool(getattr(sys, "frozen", False))
PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = PROJECT_ROOT / ".env"


def _env_search_paths() -> List[Path]:
    """
    Возвращает список мест, где может лежать .env, в порядке приоритета.

    Приоритет важен: пользователь, который положил .env рядом с .exe,
    ожидает, что именно этот файл и будет использован.
    """
    candidates: List[Path] = []

    if IS_FROZEN:
        # 1. Рядом с .exe — основной вариант для собранной программы
        candidates.append(Path(sys.executable).resolve().parent / ".env")
    else:
        # 1. Корень проекта — основной вариант для запуска из исходников
        candidates.append(PROJECT_ROOT / ".env")

    # 2. Текущая рабочая директория (запуск из другой папки)
    try:
        candidates.append(Path(os.getcwd()) / ".env")
    except Exception:
        pass

    # 3. Вверх от .exe (если .exe лежит в подпапке dist\)
    if IS_FROZEN:
        try:
            candidates.append(Path(sys.executable).resolve().parent.parent / ".env")
        except Exception:
            pass

    # Убираем дубликаты, сохраняя порядок
    unique: List[Path] = []
    for path in candidates:
        if path not in unique:
            unique.append(path)
    return unique


def find_env_file() -> Optional[Path]:
    """Возвращает первый существующий .env или None."""
    for path in _env_search_paths():
        try:
            if path.is_file():
                return path
        except OSError:
            continue
    return None


def get_env_write_path() -> Path:
    """
    Возвращает путь, КУДА сохранять .env (например, ключ из интерфейса).

    Для .exe это всегда папка рядом с .exe — то есть папка, доступная на
    запись. Внутри архива PyInstaller писать нельзя, поэтому сохранение
    «рядом с исходником» сломало бы ввод ключа.
    """
    if IS_FROZEN:
        return Path(sys.executable).resolve().parent / ".env"
    return PROJECT_ROOT / ".env"


def read_env_text(path: Path) -> str:
    """
    Читает .env, устойчиво к BOM и «неправильным» кодировкам.

    Файл мог быть создан Блокнотом (UTF-8 с BOM), bat-скриптом (CP866)
    или PowerShell-ом. Раньше BOM ломал имя первой переменной, и ключ
    молча не подхватывался — программа сообщала «ключ не задан».
    """
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp1251", "cp866"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    # Последний шанс: декодируем с заменой, лишь бы не падать
    return raw.decode("utf-8", errors="replace")


ENV_FILE: Optional[Path] = None

try:
    from dotenv import load_dotenv

    ENV_FILE = find_env_file()
    if ENV_FILE is not None:
        # override=False: реальные переменные окружения имеют приоритет над .env.
        # Строку разбираем сами — так BOM и кодировка больше не мешают.
        load_dotenv(ENV_FILE, override=False, encoding="utf-8-sig")
        # Страховка для файлов в CP866/CP1251: dotenv их не прочитает
        try:
            for line in read_env_text(ENV_FILE).splitlines():
                line = line.strip().lstrip("\ufeff")
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                if key and key not in os.environ:
                    os.environ[key] = value.strip().strip('"').strip("'")
        except Exception:
            pass
except ImportError:  # pragma: no cover — python-dotenv не установлен
    pass
except Exception:  # pragma: no cover — битый .env не должен ронять запуск
    ENV_FILE = None

# ============================================
# SSL: КОРНЕВОЙ СЕРТИФИКАТ МИНЦИФРЫ
# ============================================
# GigaChat отвечает по HTTPS, а корпоративный прокси или антивирус часто
# подменяет сертификат своим. Тогда любая проверка TLS падает с
# CERTIFICATE_VERIFY_FAILED. Правильное лечение — добавить доверенный
# корневой сертификат, а НЕ отключать проверку (verify_ssl=false).
#
# Сертификат кладётся в config/russian_trusted_root_ca.cer и подхватывается
# автоматически. Скачивать его за пользователя программа не станет:
# подмена источника сертификата — это ровно та атака, от которой он защищает.

CA_BUNDLE_FILENAME = "russian_trusted_root_ca.cer"
CA_BUNDLE_URL = "https://gu-st.ru/content/Other/doc/russian_trusted_root_ca.cer"


def _ca_bundle_dirs() -> List[Path]:
    """Папки, где программа ищет корневой сертификат Минцифры."""
    dirs: List[Path] = []
    if IS_FROZEN:
        dirs.append(Path(sys.executable).resolve().parent / "config")
    dirs.append(PROJECT_ROOT / "config")
    dirs.append(Path(os.getcwd()) / "config")
    return dirs


def find_ca_bundle() -> Optional[Path]:
    """Возвращает путь к найденному сертификату Минцифры или None."""
    for directory in _ca_bundle_dirs():
        candidate = directory / CA_BUNDLE_FILENAME
        try:
            if candidate.is_file():
                return candidate
        except OSError:
            continue
    return None


def setup_ca_bundle(logger_: Optional[Any] = None) -> Optional[str]:
    """
    Подставляет сертификат Минцифры в GIGACHAT_CA_BUNDLE_FILE, если он найден.

    Явно заданное в .env значение НЕ переопределяется — у пользователя
    может быть свой сертификат (в том числе корпоративный).

    Returns:
        Путь к используемому сертификату или None.
    """
    log = logger_ or logger

    explicit = (os.environ.get("GIGACHAT_CA_BUNDLE_FILE") or "").strip()
    if explicit:
        if os.path.isfile(explicit):
            log.info("CA-сертификат задан явно в .env: %s", explicit)
            return explicit
        log.warning(
            "GIGACHAT_CA_BUNDLE_FILE указывает на несуществующий файл: %s",
            explicit,
        )
        return None

    found = find_ca_bundle()
    if found:
        os.environ["GIGACHAT_CA_BUNDLE_FILE"] = str(found)
        log.info("CA-сертификат Минцифры найден и подключён: %s", found)
        return str(found)

    log.info(
        "CA-сертификат Минцифры не найден (ожидался в %s)",
        PROJECT_ROOT / "config" / CA_BUNDLE_FILENAME,
    )
    return None


def get_ca_bundle_status() -> Dict[str, Any]:
    """
    Полный статус сертификата Минцифры — для диагностики в интерфейсе и логе.

    Returns:
        found:          путь к подключённому сертификату или None;
        expected:       путь, где сертификат ищут в первую очередь;
        search_paths:   все места поиска (в порядке приоритета);
        explicit:       путь, заданный вручную в GIGACHAT_CA_BUNDLE_FILE;
        explicit_valid: существует ли файл, заданный вручную;
        verify_ssl:     включена ли проверка сертификата;
        url:            официальный адрес загрузки сертификата.
    """
    found = find_ca_bundle()
    explicit = (os.environ.get("GIGACHAT_CA_BUNDLE_FILE") or "").strip()
    explicit_valid = bool(explicit) and os.path.isfile(explicit)

    # Список путей без повторов: в исходниках PROJECT_ROOT и cwd — одна и та
    # же папка, и дубли только запутывают диагностику.
    search_paths: List[str] = []
    for directory in _ca_bundle_dirs():
        candidate = str(directory / CA_BUNDLE_FILENAME)
        if candidate not in search_paths:
            search_paths.append(candidate)

    return {
        "found": str(found) if found else None,
        "expected": str(_ca_bundle_dirs()[0] / CA_BUNDLE_FILENAME),
        "search_paths": search_paths,
        "explicit": explicit or None,
        "explicit_valid": explicit_valid,
        "verify_ssl": os.environ.get("GIGACHAT_VERIFY_SSL_CERTS", "true").strip().lower()
        != "false",
        "url": CA_BUNDLE_URL,
    }


def log_ssl_error_context(error: Any, logger_: Optional[Any] = None) -> None:
    """
    Пишет в лог ПОЛНЫЙ контекст SSL-ошибки, чтобы её можно было разобрать
    без повторного воспроизведения:

      * traceback (для отладки);
      * понятную причину «сертификат не найден / невалиден»;
      * все пути, где сертификат искали;
      * подсказку, где сертификат скачать.

    Ключи и токены в лог не попадают: текст ошибки проходит scrub_text().
    """
    log = logger_ or logger

    try:
        status = get_ca_bundle_status()
    except Exception as exc:  # pragma: no cover — диагностика не должна падать
        log.warning("Не удалось собрать статус CA-сертификата: %s", exc)
        return

    log.error("SSL: не удалось проверить сертификат GigaChat")

    if status["explicit"]:
        if status["explicit_valid"]:
            log.error(
                "SSL: используется сертификат из .env (GIGACHAT_CA_BUNDLE_FILE): %s",
                status["explicit"],
            )
        else:
            log.error(
                "SSL: GIGACHAT_CA_BUNDLE_FILE указывает на НЕсуществующий файл: %s",
                status["explicit"],
            )
    elif status["found"]:
        log.error(
            "SSL: сертификат Минцифры найден (%s), но проверка всё равно не прошла — "
            "скорее всего антивирус или прокси подменяет сертификат своим. "
            "Добавьте программу в исключения SSL-инспекции.",
            status["found"],
        )
    else:
        log.error(
            "SSL: корневой сертификат Минцифры НЕ НАЙДЕН — подмена TLS "
            "подтверждается только этим сертификатом"
        )

    log.error("SSL: искали сертификат в:")
    for candidate in status["search_paths"]:
        log.error("SSL:   - %s", candidate)

    log.error("SSL: скачайте сертификат: %s", status["url"])
    log.error(
        "SSL: проверка сертификата включена (verify_ssl=%s). Отключать её не нужно — "
        "она защищает от подмены трафика.",
        status["verify_ssl"],
    )
    log.error(
        "SSL: тип ошибки: %s | текст: %s",
        type(error).__name__,
        scrub_text(str(error)),
    )

    # traceback — последним: он длинный, но именно он нужен при разборе
    tb = getattr(error, "__traceback__", None)
    if tb is not None:
        log.error("SSL: traceback:", exc_info=(type(error), error, tb))
    else:
        log.error("SSL: traceback недоступен (исключение без __traceback__)")


def get_ssl_help() -> str:
    """
    Инструкция по исправлению SSL-ошибки — показывается в GUI и в логе.

    Возвращает готовый текст, чтобы интерфейс и лог не расходились.
    """
    ca_path = _ca_bundle_dirs()[0] / CA_BUNDLE_FILENAME
    return (
        "⚠️ SSL-ошибка при обращении к GigaChat.\n\n"
        "Возможные причины:\n"
        "1. Корпоративный прокси или антивирус подменяет сертификаты\n"
        "2. Не установлен корневой сертификат Минцифры\n\n"
        "Решение:\n"
        "1. Скачайте сертификат:\n"
        f"   {CA_BUNDLE_URL}\n"
        "2. Положите его в папку config рядом с программой:\n"
        f"   {ca_path}\n"
        "3. Либо пропишите свой путь в .env:\n"
        "   GIGACHAT_CA_BUNDLE_FILE=C:\\certs\\russian_trusted_root_ca.cer\n"
        "4. Перезапустите программу\n\n"
        "Отключать GIGACHAT_VERIFY_SSL_CERTS не нужно: это убирает защиту\n"
        "от подмены трафика, а проблему не решает."
    )


def log_ssl_help(logger_: Optional[Any] = None) -> None:
    """Пишет инструкцию по SSL в лог (без ключей и токенов)."""
    log = logger_ or logger
    for line in get_ssl_help().splitlines():
        log.warning("%s", line)


# ============================================
# ЛОГГЕР
# ============================================
# Используем общий логгер проекта. Если импорт не удался (например, модуль
# запущен вне структуры проекта) — падаем на стандартный logging, чтобы
# провайдер оставался полностью автономным.
#
# ВАЖНО: блок логгера обязан оставаться ВЫШЕ функций, которые пишут в лог:
# иначе имя logger ещё не существует в момент определения функции.

try:
    from logger_config import setup_logger

    logger = setup_logger(__name__)
except Exception:  # pragma: no cover
    import logging

    logger = logging.getLogger(__name__)


# ============================================
# КОНСТАНТЫ
# ============================================

DEFAULT_GIGACHAT_SCOPE = "GIGACHAT_API_PERS"
DEFAULT_GIGACHAT_MODEL = "GigaChat"
DEFAULT_GIGACHAT_TIMEOUT = 30.0
DEFAULT_OLLAMA_HOST = "http://localhost:11434"
DEFAULT_OLLAMA_MODEL = "llama3.2"
DEFAULT_OLLAMA_TIMEOUT = 20

# Значения, которые считаем «ключ не задан»
_PLACEHOLDER_VALUES = {
    "",
    "your_key_here",
    "your_key",
    "changeme",
    "none",
    "null",
    "вставьте_ключ",
    "тут_ключ",
}

# Имена исключений gigachat, означающие проблему с авторизацией/токеном.
# Только они приводят к пересозданию клиента; сетевые и SSL-ошибки — нет.
_AUTH_ERROR_NAMES = {
    "ResponseError",
    "GigaChatException",
    "AuthenticationError",
    "AuthorizationError",
}

# ============================================
# МАСКИРОВАНИЕ СЕКРЕТОВ
# ============================================


def mask_secret(value: Optional[str], keep: int = 4) -> str:
    """
    Маскирует секрет для логов.

    Возвращает "***" для пустого или короткого значения, иначе "abcd...wxyz".

    ВНИМАНИЕ: по умолчанию этот вариант НЕ используется для API-ключа —
    см. mask_credentials(). Функция оставлена для отладочных случаев,
    когда нужно сверить, ТОТ ли ключ загружен.
    """
    if not value:
        return "***"
    value = str(value)
    if len(value) <= keep * 2:
        return "***"
    return f"{value[:keep]}...{value[-keep:]}"


# Политика маскирования ключа GigaChat.
# None — показывать только "***" (значение по умолчанию, безопасный режим).
# Целое число — показывать первые/последние N символов (отладка).
# Меняйте только осознанно: любой фрагмент ключа в логе — это утечка.
GIGACHAT_MASK_KEEP: Optional[int] = None


def mask_credentials(credentials: Optional[str]) -> str:
    """
    Маска для Authorization Key GigaChat.

    По умолчанию возвращает ровно "***" — ключ не попадает в лог ни одним
    символом. Возвращает "***" и для пустого значения, чтобы по логу нельзя
    было отличить «ключ не задан» от «ключ задан».
    """
    if GIGACHAT_MASK_KEEP is None:
        return "***"
    return mask_secret(credentials, keep=GIGACHAT_MASK_KEEP)


def scrub_text(text: str) -> str:
    """
    Вычищает секреты из произвольного текста (например, текста исключения).

    Нужна, потому что некоторые библиотеки и HTTP-клиенты любят включать
    заголовки/URL с токеном в сообщение об ошибке.
    """
    if not text:
        return ""
    scrubbed = str(text)

    # Прямые значения секретов из окружения
    for name in ("GIGACHAT_CREDENTIALS", "GIGACHAT_PASSWORD"):
        secret = os.environ.get(name)
        if secret and len(secret) >= 6:
            scrubbed = scrubbed.replace(secret, "***")

    # Общие шаблоны: Bearer <token>, Authorization: Basic <base64>, access_token=<...>
    scrubbed = re.sub(r"(?i)(bearer\s+)[A-Za-z0-9._\-+/=]{8,}", r"\1***", scrubbed)
    scrubbed = re.sub(r"(?i)(basic\s+)[A-Za-z0-9+/=]{8,}", r"\1***", scrubbed)
    scrubbed = re.sub(
        r"(?i)(access_token[\"'\s:=]+)[A-Za-z0-9._\-+/=]{8,}", r"\1***", scrubbed
    )
    scrubbed = re.sub(
        r"(?i)(authorization[\"'\s:=]+)[A-Za-z0-9._\-+/=]{8,}", r"\1***", scrubbed
    )

    return scrubbed


# ============================================
# GIGACHAT — КОНФИГУРАЦИЯ
# ============================================


def get_gigachat_credentials() -> str:
    """Возвращает Authorization Key из окружения (после загрузки .env)."""
    return (os.environ.get("GIGACHAT_CREDENTIALS") or "").strip()


def is_gigachat_configured() -> bool:
    """
    Проверяет, задан ли ключ GigaChat.

    Сама библиотека gigachat при этом НЕ требуется — проверка чисто по .env,
    поэтому функция безопасна для вызова до установки зависимостей.
    """
    return get_gigachat_credentials().lower() not in _PLACEHOLDER_VALUES


# ============================================
# GIGACHAT — ЛЕНИВЫЙ SINGLETON КЛИЕНТ
# ============================================

_gigachat_client: Optional[Any] = None
_gigachat_config_key: Optional[tuple] = None
_gigachat_client_timeout: Optional[float] = None


def _current_config_key() -> tuple:
    """Отпечаток конфигурации: при её смене клиент пересоздаётся."""
    return (
        get_gigachat_credentials(),
        os.environ.get("GIGACHAT_SCOPE", DEFAULT_GIGACHAT_SCOPE),
        os.environ.get("GIGACHAT_MODEL", DEFAULT_GIGACHAT_MODEL),
        os.environ.get("GIGACHAT_VERIFY_SSL_CERTS", "true"),
        os.environ.get("GIGACHAT_CA_BUNDLE_FILE", ""),
    )


def reset_gigachat_client() -> None:
    """
    Сбрасывает кэшированный клиент (например, после смены ключа без перезапуска
    программы). Старое соединение закрывается, ошибки игнорируются.
    """
    global _gigachat_client, _gigachat_config_key, _gigachat_client_timeout

    if _gigachat_client is not None:
        try:
            _gigachat_client.close()
        except Exception:
            pass

    _gigachat_client = None
    _gigachat_config_key = None
    _gigachat_client_timeout = None
    logger.info("Клиент GigaChat сброшен")


def _resolve_timeout(timeout: Optional[float] = None) -> float:
    """
    Возвращает таймаут клиента в секундах.

    Приоритет: аргумент вызова → GIGACHAT_TIMEOUT из .env → значение
    по умолчанию. Некорректное значение в .env не ломает запрос, а
    откатывается к значению по умолчанию (с предупреждением в лог).
    """
    if timeout is not None:
        try:
            value = float(timeout)
            if value > 0:
                return value
            logger.warning(
                "Некорректный timeout=%r — использую значение по умолчанию", timeout
            )
        except (TypeError, ValueError):
            logger.warning(
                "Некорректный timeout=%r — использую значение по умолчанию", timeout
            )

    try:
        return float(os.environ.get("GIGACHAT_TIMEOUT", DEFAULT_GIGACHAT_TIMEOUT))
    except (TypeError, ValueError):
        logger.warning("Некорректный GIGACHAT_TIMEOUT — использую значение по умолчанию")
        return DEFAULT_GIGACHAT_TIMEOUT


def _get_gigachat_client(timeout: Optional[float] = None) -> Optional[Any]:
    """
    Возвращает клиент GigaChat, создавая его лениво при первом вызове.

    Ленивость важна: отсутствие ключа или библиотеки не должно ломать
    программу на старте — GUI должен открыться и показать статус.

    timeout — необязательный таймаут конкретного вызова (секунды). Такие
    клиенты по умолчанию НЕ кэшируются: клиент с «чужим» таймаутом,
    оставшийся в кэше, менял бы поведение следующих запросов. Исключение —
    успешный вызов, после которого клиент становится основным.
    """
    global _gigachat_client, _gigachat_config_key, _gigachat_client_timeout

    if not is_gigachat_configured():
        logger.warning("GigaChat не настроен: GIGACHAT_CREDENTIALS не задан в .env")
        return None

    config_key = _current_config_key()
    effective_timeout = _resolve_timeout(timeout)

    # Клиент уже создан, конфигурация не менялась и таймаут тот же — переиспользуем
    if _gigachat_client is not None and _gigachat_config_key == config_key:
        if _gigachat_client_timeout is None or _gigachat_client_timeout == effective_timeout:
            return _gigachat_client

    # Конфигурация изменилась — пересоздаём
    if _gigachat_client is not None and _gigachat_config_key != config_key:
        logger.info("Конфигурация GigaChat изменилась — пересоздаю клиент")
        reset_gigachat_client()

    try:
        from gigachat import GigaChat
    except ImportError:
        logger.error(
            "Библиотека gigachat не установлена. "
            "Установите: pip install gigachat python-dotenv"
        )
        return None

    credentials = get_gigachat_credentials()
    scope = os.environ.get("GIGACHAT_SCOPE", DEFAULT_GIGACHAT_SCOPE).strip()
    model = os.environ.get("GIGACHAT_MODEL", DEFAULT_GIGACHAT_MODEL).strip()

    timeout = effective_timeout

    # SSL-верификация включена, если явно не отключена ("false")
    verify_ssl = (
        os.environ.get("GIGACHAT_VERIFY_SSL_CERTS", "true").strip().lower() != "false"
    )

    # Сертификат Минцифры подхватывается автоматически: без него в сетях с
    # корпоративной подменой TLS любой запрос падает с CERTIFICATE_VERIFY_FAILED.
    setup_ca_bundle()
    ca_bundle = (os.environ.get("GIGACHAT_CA_BUNDLE_FILE") or "").strip() or None

    base_url = (os.environ.get("GIGACHAT_BASE_URL") or "").strip() or None

    client_kwargs: Dict[str, Any] = {
        "credentials": credentials,
        "scope": scope,
        "model": model,
        "timeout": timeout,
        "verify_ssl_certs": verify_ssl,
        "ca_bundle_file": ca_bundle,
    }
    if base_url:
        # Позволяет работать через прокси, если прямой доступ закрыт
        client_kwargs["base_url"] = base_url

    try:
        client = GigaChat(**client_kwargs)

        # Клиент с нестандартным таймаутом оставляем вызывающему коду:
        # в общий кэш он не попадает, если это разовый вызов.
        if timeout == _resolve_timeout():
            _gigachat_client = client
            _gigachat_config_key = config_key
            _gigachat_client_timeout = timeout

        # ВАЖНО: сам ключ не логируем — только маску
        logger.info(
            "Клиент GigaChat создан | credentials=%s | scope=%s | model=%s | "
            "timeout=%sс | verify_ssl=%s | ca_bundle=%s",
            mask_credentials(credentials),
            scope,
            model,
            timeout,
            verify_ssl,
            ca_bundle or "системный",
        )
        return client

    except Exception as e:
        logger.error("Не удалось создать клиент GigaChat: %s", scrub_text(str(e)))
        _gigachat_client = None
        _gigachat_config_key = None
        _gigachat_client_timeout = None
        return None


def warmup_gigachat() -> bool:
    """
    Прогревает авторизацию GigaChat (получает токен заранее).

    Вызывать в фоновом потоке на старте GUI — тогда первый реальный запрос
    не будет ждать OAuth-хендшейк. Возвращает True при успехе.
    """
    client = _get_gigachat_client()
    if client is None:
        return False

    try:
        start = time.time()
        client.get_token()
        clear_ssl_error()
        logger.info("Прогрев GigaChat выполнен за %.2f сек", time.time() - start)
        return True
    except Exception as e:
        if is_ssl_error(e):
            note_ssl_error(e)
        logger.warning("Прогрев GigaChat не удался: %s", scrub_text(str(e)))
        return False


# ============================================
# SSL: ОБНАРУЖЕНИЕ ПРОБЛЕМЫ
# ============================================
# Последняя SSL-ошибка хранится, чтобы интерфейс мог показать красную
# плашку «см. README», а не молча уходить в шаблонный текст.

_last_ssl_error: Optional[str] = None
_ssl_help_logged = False


def is_ssl_error(error: Any) -> bool:
    """
    Определяет, что ошибка связана с проверкой TLS-сертификата.

    Проверяем и тип, и текст: разные версии httpx/ssl приносят это
    то как ConnectError, то как SSLError, то как ошибку внутри цепочки.
    """
    text = scrub_text(str(error)).lower()
    markers = (
        "certificate_verify_failed",
        "certificate verify failed",
        "self-signed certificate",
        "self signed certificate",
        "ssl: ",
        "sslerror",
        "unknown ca",
    )
    if any(marker in text for marker in markers):
        return True
    return type(error).__name__ in {"SSLError", "SSLCertVerificationError"}


def note_ssl_error(error: Any) -> None:
    """Запоминает SSL-ошибку и один раз печатает инструкцию в лог."""
    global _last_ssl_error, _ssl_help_logged

    _last_ssl_error = scrub_text(str(error))

    if not _ssl_help_logged:
        _ssl_help_logged = True
        logger.error("Обнаружена SSL-ошибка при обращении к GigaChat")
        log_ssl_help()


def get_last_ssl_error() -> Optional[str]:
    """Текст последней SSL-ошибки или None, если её не было."""
    return _last_ssl_error


def has_ssl_error() -> bool:
    """Была ли в этом запуске SSL-ошибка (для статуса в интерфейсе)."""
    return _last_ssl_error is not None


def clear_ssl_error() -> None:
    """Сбрасывает отметку SSL-ошибки (после успешного запроса)."""
    global _last_ssl_error
    _last_ssl_error = None


# ============================================
# GIGACHAT — ГЕНЕРАЦИЯ
# ============================================


def _extract_message_content(response: Any) -> str:
    """Достаёт текст ответа из любого поддерживаемого формата ответа."""
    if response is None:
        return ""

    # Основной путь: response.choices[0].message.content
    choices = getattr(response, "choices", None)
    if choices:
        message = getattr(choices[0], "message", None)
        content = getattr(message, "content", None)
        if content:
            return str(content).strip()
        # На случай, если пришёл "сырой" dict
        if isinstance(choices[0], dict):
            return str(choices[0].get("message", {}).get("content", "")).strip()

    # Фолбэк для dict-подобных ответов
    if isinstance(response, dict):
        try:
            return str(response["choices"][0]["message"]["content"]).strip()
        except (KeyError, IndexError, TypeError):
            return ""

    return ""


def _generate_gigachat_streaming(
    client: Any, system_prompt: str, prompt: str, temperature: float, max_tokens: int
) -> str:
    """
    Потоковая генерация: текст приходит по мере готовности.

    Это быстрее по ощущениям (и экономит время на разбор полного ответа),
    поэтому используется как основной путь.
    """
    from gigachat.models import Chat, Messages, MessagesRole

    payload = Chat(
        messages=[
            Messages(role=MessagesRole.SYSTEM, content=system_prompt),
            Messages(role=MessagesRole.USER, content=prompt),
        ],
        temperature=temperature,
        max_tokens=max_tokens,
        stream=True,
    )

    chunks: List[str] = []
    for chunk in client.stream(payload):
        content = _extract_message_content(chunk)
        if content:
            chunks.append(content)

    return "".join(chunks).strip()


def _generate_gigachat_sync(
    client: Any, system_prompt: str, prompt: str, temperature: float, max_tokens: int
) -> str:
    """Обычная (непотоковая) генерация — фолбэк, если стриминг не сработал."""
    from gigachat.models import Chat, Messages, MessagesRole

    payload = Chat(
        messages=[
            Messages(role=MessagesRole.SYSTEM, content=system_prompt),
            Messages(role=MessagesRole.USER, content=prompt),
        ],
        temperature=temperature,
        max_tokens=max_tokens,
    )

    return _extract_message_content(client.chat(payload))


def generate_gigachat(
    prompt: str,
    system_prompt: Optional[str] = None,
    temperature: float = 0.85,
    max_tokens: int = 280,
    use_stream: bool = True,
    timeout: Optional[float] = None,
) -> Optional[str]:
    """
    Генерирует текст через GigaChat.

    Args:
        prompt:        пользовательский запрос
        system_prompt: системная роль (по умолчанию — копирайтер по автоперевозкам)
        temperature:   креативность (0.0–2.0)
        max_tokens:    ограничение длины ответа (меньше = быстрее)
        use_stream:    потоковая генерация (быстрее отдаёт первый токен)
        timeout:       таймаут запроса в секундах; None — взять GIGACHAT_TIMEOUT
                       из .env (по умолчанию 30 с). Запрос ВСЕГДА ограничен
                       таймаутом: «зависнуть» на недоступной сети программа
                       не должна.

    Returns:
        Текст ответа или None при любой ошибке. Исключения не выбрасываются.
    """
    if not prompt or not str(prompt).strip():
        logger.warning("generate_gigachat: пустой prompt")
        return None

    if system_prompt is None:
        system_prompt = DEFAULT_SYSTEM_PROMPT

    if not is_gigachat_configured():
        logger.warning("generate_gigachat: ключ GigaChat не задан (см. .env)")
        return None

    client = _get_gigachat_client(timeout)
    if client is None:
        return None

    try:
        temperature = float(temperature)
    except (TypeError, ValueError):
        temperature = 0.85

    try:
        max_tokens = int(max_tokens)
    except (TypeError, ValueError):
        max_tokens = 280

    start = time.time()

    try:
        text = ""
        if use_stream:
            try:
                text = _generate_gigachat_streaming(
                    client, system_prompt, str(prompt), temperature, max_tokens
                )
            except Exception as stream_error:
                # Стриминг может не поддерживаться прокси/сетью — идём обычным путём
                logger.warning(
                    "Потоковая генерация GigaChat не удалась (%s), "
                    "переключаюсь на обычный режим",
                    scrub_text(str(stream_error)),
                )
                text = _generate_gigachat_sync(
                    client, system_prompt, str(prompt), temperature, max_tokens
                )
        else:
            text = _generate_gigachat_sync(
                client, system_prompt, str(prompt), temperature, max_tokens
            )

        elapsed = time.time() - start

        if not text:
            logger.warning("GigaChat вернул пустой ответ за %.2f сек", elapsed)
            return None

        logger.info(
            "GigaChat: ответ получен за %.2f сек (символов: %d)", elapsed, len(text)
        )
        return text

    except Exception as e:
        elapsed = time.time() - start
        error_name = type(e).__name__
        logger.error(
            "Ошибка GigaChat за %.2f сек: %s: %s",
            elapsed,
            error_name,
            scrub_text(str(e)),
        )

        # Сбрасываем клиент ТОЛЬКО при ошибке авторизации (протух токен,
        # неверный ключ). Сетевые/SSL-ошибки клиент не портят — пересоздавать
        # его нельзя, иначе теряется прогрев и TLS-соединение.
        if error_name in _AUTH_ERROR_NAMES or "401" in scrub_text(str(e)):
            logger.info("Сбрасываю клиент GigaChat после ошибки авторизации")
            reset_gigachat_client()

        # SSL — отдельная причина: дело не в ключе и не в сети, а в проверке
        # сертификата. Пишем в лог полный контекст (traceback, пути поиска
        # сертификата, ссылку на скачивание) и запоминаем, чтобы интерфейс
        # показал плашку «как исправить».
        if is_ssl_error(e):
            note_ssl_error(e)
            log_ssl_error_context(e)

        return None


DEFAULT_SYSTEM_PROMPT = (
    "Ты — опытный копирайтер транспортной компании. "
    "Пиши короткие живые коммерческие предложения по автоперевозкам на русском языке. "
    "Стиль энергичный, как в чате перевозчиков, но без ошибок и лишней воды. "
    "Не выдумывай цены, названия компаний и сроки. Не добавляй подписи и обращения."
)


# ============================================
# OLLAMA — ПРОВЕРКА И ГЕНЕРАЦИЯ
# ============================================


def _ollama_host() -> str:
    return os.environ.get("OLLAMA_HOST", DEFAULT_OLLAMA_HOST).rstrip("/")


def get_ollama_models(timeout: float = 2.0) -> List[str]:
    """Возвращает список доступных моделей Ollama (пустой список при ошибке)."""
    try:
        import requests

        response = requests.get(f"{_ollama_host()}/api/tags", timeout=timeout)
        if response.status_code == 200:
            return [m.get("name", "") for m in response.json().get("models", [])]
        return []
    except Exception:
        return []


def is_ollama_available(timeout: float = 2.0) -> bool:
    """Проверяет, запущен ли Ollama и есть ли хотя бы одна модель."""
    models = get_ollama_models(timeout=timeout)
    if models:
        return True

    # Сервер может быть запущен, но без моделей — проверим сам факт доступности
    try:
        import requests

        response = requests.get(f"{_ollama_host()}/api/tags", timeout=timeout)
        return response.status_code == 200 and bool(response.json().get("models"))
    except Exception:
        return False


def generate_ollama(
    prompt: str,
    model: Optional[str] = None,
    temperature: float = 0.8,
    timeout: int = DEFAULT_OLLAMA_TIMEOUT,
    max_tokens: int = 300,
    system_prompt: Optional[str] = None,
) -> Optional[str]:
    """
    Генерирует текст через локальный Ollama.

    Returns:
        Текст ответа или None при любой ошибке. Исключения не выбрасываются.
    """
    if not prompt or not str(prompt).strip():
        logger.warning("generate_ollama: пустой prompt")
        return None

    try:
        import requests
    except ImportError:
        logger.error("Библиотека requests не установлена — Ollama недоступен")
        return None

    models = get_ollama_models()
    if not models:
        logger.warning("generate_ollama: Ollama недоступен или нет моделей")
        return None

    if not model or model not in models:
        model = models[0]
        logger.info("Автоматически выбрана модель Ollama: %s", model)

    try:
        temperature = float(temperature)
    except (TypeError, ValueError):
        temperature = 0.8

    payload: Dict[str, Any] = {
        "model": model,
        "prompt": str(prompt),
        "stream": False,
        "options": {
            "temperature": temperature,
            "top_p": 0.9,
            "top_k": 50,
            "num_predict": int(max_tokens),
            "repeat_penalty": 1.2,
        },
    }
    if system_prompt:
        payload["system"] = system_prompt

    start = time.time()
    try:
        response = requests.post(
            f"{_ollama_host()}/api/generate", json=payload, timeout=timeout
        )
        elapsed = time.time() - start
        logger.info(
            "Ollama: ответ за %.2f сек (HTTP %s, модель %s)",
            elapsed,
            response.status_code,
            model,
        )

        if response.status_code != 200:
            logger.error("Ollama вернул HTTP %s", response.status_code)
            return None

        text = (response.json().get("response") or "").strip()
        if not text:
            logger.warning("Ollama вернул пустой ответ")
            return None

        return text

    except requests.Timeout:
        logger.error("Таймаут запроса к Ollama (%s сек)", timeout)
        return None
    except Exception as e:
        logger.error("Ошибка запроса к Ollama: %s", scrub_text(str(e)))
        return None


# ============================================
# ЕДИНАЯ ТОЧКА ВХОДА
# ============================================


def get_available_provider() -> str:
    """
    Определяет лучший доступный провайдер.

    Returns:
        "gigachat" | "ollama" | "none"
    """
    if is_gigachat_configured():
        return "gigachat"
    if is_ollama_available():
        return "ollama"
    return "none"


def generate_text(
    prompt: str,
    provider: str = "gigachat",
    system_prompt: Optional[str] = None,
    temperature: float = 0.85,
    max_tokens: int = 280,
    ollama_model: Optional[str] = None,
    ollama_timeout: int = DEFAULT_OLLAMA_TIMEOUT,
) -> Optional[str]:
    """
    Единая точка генерации текста.

    provider="gigachat" — GigaChat, при неудаче автоматический откат на Ollama
    provider="ollama"   — только Ollama
    provider="auto"     — GigaChat, если настроен, иначе Ollama

    Returns:
        Текст или None, если оба провайдера не сработали.
    """
    if provider == "auto":
        provider = get_available_provider()

    if provider == "gigachat":
        result = generate_gigachat(
            prompt=prompt,
            system_prompt=system_prompt,
            temperature=temperature,
            max_tokens=max_tokens,
        )
        if result:
            return result

        logger.info("GigaChat не дал результат — пробую Ollama как запасной вариант")
        return generate_ollama(
            prompt=prompt,
            model=ollama_model,
            temperature=temperature,
            timeout=ollama_timeout,
            max_tokens=max_tokens,
            system_prompt=system_prompt,
        )

    if provider == "ollama":
        return generate_ollama(
            prompt=prompt,
            model=ollama_model,
            temperature=temperature,
            timeout=ollama_timeout,
            max_tokens=max_tokens,
            system_prompt=system_prompt,
        )

    logger.warning("generate_text: нет доступного провайдера (provider=%s)", provider)
    return None


def get_status_info() -> Dict[str, Any]:
    """Краткая сводка о состоянии провайдеров — для логов и UI."""
    try:
        model = os.environ.get("GIGACHAT_MODEL", DEFAULT_GIGACHAT_MODEL)
    except Exception:
        model = DEFAULT_GIGACHAT_MODEL

    ca_bundle = (os.environ.get("GIGACHAT_CA_BUNDLE_FILE") or "").strip()
    found_ca = find_ca_bundle()
    env_file = ENV_FILE or find_env_file()

    return {
        "gigachat_configured": is_gigachat_configured(),
        "gigachat_model": model,
        "gigachat_credentials": mask_credentials(get_gigachat_credentials()),
        "ollama_available": is_ollama_available(),
        "ollama_models": get_ollama_models(),
        "env_path": str(env_file) if env_file else str(get_env_write_path()),
        "env_exists": bool(env_file and env_file.exists()),
        "env_found_paths": [str(p) for p in _env_search_paths()],
        "ca_bundle": ca_bundle,
        "ca_bundle_found": str(found_ca) if found_ca else None,
        "ca_bundle_expected": str(_ca_bundle_dirs()[0] / CA_BUNDLE_FILENAME),
        "ssl_error": has_ssl_error(),
        "frozen": IS_FROZEN,
    }


# ============================================
# САМОПРОВЕРКА
# ============================================

# Сертификат Минцифры подхватывается сразу при импорте: так к моменту
# первого запроса GIGACHAT_CA_BUNDLE_FILE уже заполнен.
try:
    setup_ca_bundle()
except Exception:  # pragma: no cover — диагностика не должна мешать работе
    pass


if __name__ == "__main__":
    # Консоль Windows часто в cp1251 — переключаем вывод в UTF-8,
    # чтобы русский текст и эмодзи печатались, а не падали с UnicodeEncodeError
    try:
        import sys

        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    logger.info("=" * 70)
    logger.info("САМОПРОВЕРКА llm_provider")
    logger.info("=" * 70)

    info = get_status_info()
    print(f"Запуск из EXE:        {'да' if info['frozen'] else 'нет (исходники)'}")
    print(f"Файл .env:            {info['env_path']} (существует: {info['env_exists']})")
    print(f"GigaChat настроен:    {'✅ да' if info['gigachat_configured'] else '❌ нет'}")
    print(f"GigaChat credentials: {info['gigachat_credentials']}")
    print(f"GigaChat модель:      {info['gigachat_model']}")
    print(f"Ollama доступен:      {'✅ да' if info['ollama_available'] else '❌ нет'}")
    print(f"Ollama модели:        {', '.join(info['ollama_models']) or '—'}")
    print(f"Активный провайдер:   {get_available_provider()}")
    print()
    print(f"Сертификат Минцифры:  {info['ca_bundle'] or '— не задан —'}")
    print(f"  найден в проекте:   {info['ca_bundle_found'] or '❌ нет'}")
    print(f"  ожидаемый путь:     {info['ca_bundle_expected']}")
    print(f"SSL-ошибка в сессии:  {'⚠️ да' if info['ssl_error'] else 'нет'}")
    print()
    print("Искали .env в:")
    for candidate in info["env_found_paths"]:
        print(f"  - {candidate}")
