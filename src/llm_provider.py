#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Универсальный провайдер LLM для «Генератора КП».

Бэкенд один:
  * GigaChat (облачный, Сбер) — ключ берётся ТОЛЬКО из .env

Локальный Ollama из проекта удалён: он требовал отдельной установки на каждой
машине, а программа собирается в EXE и работает у перевозчиков «из коробки».

Принципы безопасности:
  * API-ключ никогда не логируется: любые упоминания маскируются как "***"
  * Ключ не хранится в коде — только в .env (который в .gitignore)
  * SSL-верификация включена по умолчанию
  * Ни одна функция не выбрасывает исключение наружу: при ошибке возвращается None

Принципы скорости:
  * Клиент GigaChat создаётся лениво и кэшируется (singleton) — токен и
    TLS-соединение переиспользуются между запросами
  * Генерация идёт ОБЫЧНЫМ (непотоковым) запросом: стриминг в сетях с
    SSL-инспекцией (Kaspersky, Pro32) обрывается, поэтому он отключён по
    умолчанию и включается только явным use_stream=True
  * Короткий промпт и небольшой max_tokens — меньше текста, быстрее ответ
  * Есть отдельная функция warmup() для прогрева авторизации в фоне
"""

import hashlib
import os
import ssl
import sys
import re
import threading
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

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

    # 4. Домашняя папка пользователя — самый последний запасной вариант.
    # Полезно, если программа запускается из нестандартного места, а ключ
    # пользователь положил в профиль. Приоритет ниже всех остальных.
    try:
        candidates.append(Path.home() / ".env")
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

# Имя файла-склейки, который программа собирает сама (см. build_ca_bundle).
MERGED_CA_BUNDLE_FILENAME = "ca_bundle_merged.pem"


def _config_dirs() -> List[Path]:
    """
    Папки config\\, где может лежать сертификат Минцифры, по приоритету.

    Порядок: рядом с .exe (сборка onedir) → корень проекта (запуск из
    исходников) → текущая папка. Дубликаты убираются.
    """
    dirs: List[Path] = []

    if IS_FROZEN:
        try:
            dirs.append(Path(sys.executable).resolve().parent / "config")
        except Exception:
            pass

    dirs.append(PROJECT_ROOT / "config")

    try:
        dirs.append(Path(os.getcwd()) / "config")
    except Exception:
        pass

    unique: List[Path] = []
    for directory in dirs:
        if directory not in unique:
            unique.append(directory)
    return unique


def _ca_bundle_search_paths() -> List[Path]:
    """
    ВСЕ места, где программа ищет сертификат Минцифры, в порядке приоритета.

    Раньше проверялась только папка config\\ — и пользователь, который
    положил сертификат просто в корень проекта (или рядом с .exe), видел
    «сертификат не найден», хотя файл лежал на виду. Теперь проверяются
    все разумные места, а диагностика показывает полный список.
    """
    paths: List[Path] = []

    def add(path: Optional[Path]) -> None:
        """Добавляет путь без дубликатов."""
        if path is None:
            return
        try:
            resolved = Path(path)
        except Exception:
            return
        if resolved not in paths:
            paths.append(resolved)

    # 1. Все папки config\: и ca.cer, и готовая склейка внутри них
    for directory in _config_dirs():
        add(directory / CA_BUNDLE_FILENAME)
        add(directory / MERGED_CA_BUNDLE_FILENAME)

    # 2. Прямо рядом с .exe (сборка onedir: .exe лежит в корне папки сборки)
    if IS_FROZEN:
        try:
            exe_dir = Path(sys.executable).resolve().parent
            add(exe_dir / CA_BUNDLE_FILENAME)
            add(exe_dir / MERGED_CA_BUNDLE_FILENAME)
        except Exception:
            pass

    # 3. Корень проекта — «просто положил файл в папку проекта»
    add(PROJECT_ROOT / CA_BUNDLE_FILENAME)
    add(PROJECT_ROOT / MERGED_CA_BUNDLE_FILENAME)

    # 4. Текущая рабочая папка (программу запустили из другого места)
    try:
        add(Path(os.getcwd()) / CA_BUNDLE_FILENAME)
        add(Path(os.getcwd()) / MERGED_CA_BUNDLE_FILENAME)
    except Exception:
        pass

    # 5. Домашняя папка — последний запасной вариант
    try:
        home = Path.home()
        add(home / ".config" / "gigachat" / CA_BUNDLE_FILENAME)
        add(home / CA_BUNDLE_FILENAME)
    except Exception:
        pass

    return paths


def _ca_bundle_dirs() -> List[Path]:
    """
    Папки, где программа ищет корневой сертификат Минцифры.

    Оставлена для совместимости (ею пользуются get_ssl_help и тесты);
    фактический поиск идёт по полному списку _ca_bundle_search_paths().
    """
    return _config_dirs()


def find_ca_bundle() -> Optional[Path]:
    """
    Возвращает путь к найденному сертификату Минцифры или None.

    Проверяются все места из _ca_bundle_search_paths() — включая корень
    проекта и папку рядом с .exe, а не только config\\.
    """
    for candidate in _ca_bundle_search_paths():
        try:
            if candidate.is_file():
                return candidate
        except OSError:
            continue
    return None


def get_ca_bundle_problem() -> Optional[str]:
    """
    Возвращает понятный текст проблемы с сертификатом или None.

    Нужна для диагностики: если файл лежит, но пустой или слишком мал,
    формально он «найден», а на деле не работает. Лучше сказать об этом
    сразу, чем показывать «✅ найден» и потом ловить SSL-ошибку.
    """
    explicit = (os.environ.get("GIGACHAT_CA_BUNDLE_FILE") or "").strip()
    if explicit and not os.path.isfile(explicit):
        return f"путь из .env (GIGACHAT_CA_BUNDLE_FILE) не существует: {explicit}"

    found = find_ca_bundle()
    if found is None:
        return None  # «не найден» — это отдельный сценарий, см. get_ca_bundle_status

    try:
        size = found.stat().st_size
    except OSError as exc:
        return f"не удалось прочитать файл сертификата: {exc}"

    if size < 256:
        return f"файл сертификата подозрительно мал ({size} байт): {found}"

    try:
        text = read_env_text(found)
    except Exception as exc:
        return f"файл сертификата не читается: {exc}"

    if "BEGIN CERTIFICATE" not in text:
        return f"в файле нет сертификата (ожидался PEM/Base64): {found}"

    return None


def _cert_common_name(der: bytes) -> str:
    """
    Возвращает Common Name сертификата (для диагностики) или пустую строку.

    Используется библиотека cryptography (она и так стоит вместе с gigachat).
    Если её нет — функция молча возвращает пустую строку: диагностика не
    должна ломаться из-за отсутствия необязательной зависимости.
    """
    try:
        from cryptography import x509  # type: ignore

        cert = x509.load_der_x509_certificate(der)
        for attribute in cert.subject:
            if attribute.oid._name in ("commonName", "CN"):
                return str(attribute.value)
    except Exception:
        return ""
    return ""


def _extra_ca_files() -> List[Path]:
    """
    Дополнительные сертификаты из папок config\\ — все *.cer/*.crt/*.pem.

    Зачем: если ключ задан явно (GIGACHAT_CA_BUNDLE_FILE) или у пользователя
    лежит ещё и корень антивируса, они нужны В ОДНОЙ склейке — цепочка
    может требовать сразу оба.
    """
    files: List[Path] = []
    for directory in _config_dirs():
        try:
            entries = sorted(directory.iterdir())
        except OSError:
            continue
        for entry in entries:
            try:
                if not entry.is_file():
                    continue
            except OSError:
                continue
            if entry.suffix.lower() not in (".cer", ".crt", ".pem"):
                continue
            if entry.name == MERGED_CA_BUNDLE_FILENAME:
                continue  # это наша же склейка, второй раз не нужна
            if entry not in files:
                files.append(entry)
    return files


def _certifi_bundle() -> Optional[Path]:
    """
    Путь к штатному набору корневых сертификатов (cacert.pem от certifi).

    Нужен потому, что httpx (а значит и GigaChat) по умолчанию проверяет
    именно этот файл. Стоит задать свой — и все «обычные» сайты перестанут
    проверяться, поэтому склейка обязана включать и certifi тоже.
    """
    try:
        import certifi  # type: ignore

        path = Path(certifi.where())
        if path.is_file():
            return path
    except Exception:
        pass
    return None


def _is_path_inside(path: Path, directory: Path) -> bool:
    """True, если path лежит внутри directory (без исключений наружу)."""
    try:
        path.resolve().relative_to(directory.resolve())
        return True
    except Exception:
        return False


_WORK_DIR_CACHE: Dict[str, Optional[Path]] = {"path": None}


def get_work_dir() -> Path:
    """
    Рабочая папка для служебных файлов программы (временные сертификаты и т.п.).

    Почему не просто tempfile.gettempdir(): внутри собранного .exe системная
    временная папка может определиться как папка самой программы — и тогда
    служебные файлы сыпались рядом с .exe. Это выглядит как мусор, а на
    машине, где папка программы доступна только для чтения, программа
    вообще не смогла бы собрать набор сертификатов.

    Порядок выбора:
      1. системная временная папка, но только если это НЕ папка программы
         (иначе получится ровно тот мусор, от которого уходим);
      2. подпапка runtime рядом с программой;
      3. подпапка программы (последний вариант — работать важнее, чем чистота).
    """
    if _WORK_DIR_CACHE["path"] is not None:
        return _WORK_DIR_CACHE["path"]

    import tempfile

    def usable(path: Path) -> bool:
        """Папка пригодна, если в неё реально можно записать файл."""
        try:
            path.mkdir(parents=True, exist_ok=True)
            probe = path / ".write_probe"
            probe.write_text("ok", encoding="utf-8")
            probe.unlink()
            return True
        except Exception:
            return False

    def is_app_dir(path: Path) -> bool:
        """True, если это сама папка программы (или лежит внутри неё)."""
        return _is_path_inside(path, PROJECT_ROOT)

    candidates: List[Path] = []

    system_temp: Optional[Path] = None
    try:
        system_temp = Path(tempfile.gettempdir())
    except Exception:
        system_temp = None

    # Системная временная папка — обычный и самый правильный вариант
    if system_temp is not None and not is_app_dir(system_temp):
        candidates.append(system_temp)
    # Внутри неё — своя подпапка: на случай, если корень временной папки
    # недоступен для записи (бывает при жёстких политиках).
    if system_temp is not None and not is_app_dir(system_temp):
        candidates.append(system_temp / "GeneratorKP")
    # Рядом с программой — только в подпапку, чтобы не мусорить в корне
    try:
        candidates.append(PROJECT_ROOT / "runtime")
    except Exception:
        pass
    candidates.append(PROJECT_ROOT)

    for candidate in candidates:
        if usable(candidate):
            _WORK_DIR_CACHE["path"] = candidate
            logger.debug("Служебная папка программы: %s", candidate)
            return candidate

    _WORK_DIR_CACHE["path"] = PROJECT_ROOT
    logger.debug("Служебная папка программы (запасной вариант): %s", PROJECT_ROOT)
    return PROJECT_ROOT


def make_work_subdir(name: str) -> Path:
    """
    Возвращает подпапку рабочей папки с указанным именем и переносит туда
    файлы, оставшиеся от прошлых версий программы в корне папки программы.

    Нужна для чистоты: раньше служебные файлы могли попасть прямо в папку
    программы, и они там так и оставались.
    """
    directory = get_work_dir() / name
    try:
        directory.mkdir(parents=True, exist_ok=True)
    except Exception:
        return directory

    # Переносим возможные «старые» файлы из корня папки программы
    try:
        legacy_files = [PROJECT_ROOT / name]
        if name == "gigachat_roots":
            legacy_files.append(PROJECT_ROOT / "gigachat_windows_roots.txt")
        for legacy in legacy_files:
            if legacy == directory or not legacy.exists():
                continue
            if legacy.is_dir():
                for item in legacy.iterdir():
                    try:
                        item.replace(directory / item.name)
                    except Exception:
                        pass
                try:
                    legacy.rmdir()
                except OSError:
                    pass
            else:
                try:
                    legacy.replace(directory / legacy.name)
                except Exception:
                    pass
    except Exception:
        pass

    return directory


def _copy_windows_cert_to_pem(cert_id: str, pem_text: str) -> Optional[Path]:
    """
    Кладёт PEM-текст сертификата в стабильный файл в рабочей папке.

    Имя файла включает отпечаток, поэтому файл создаётся один раз и потом
    просто переиспользуется. Возвращает путь или None.
    """
    safe = re.sub(r"[^0-9A-Fa-f]", "", cert_id)[:40] or hashlib.sha1(
        pem_text.encode("utf-8")
    ).hexdigest()[:16]
    target = get_work_dir() / f"gigachat_extra_root_{safe.upper()}.pem"
    try:
        if target.is_file() and target.read_text(encoding="utf-8", errors="ignore") == pem_text:
            return target
        target.write_text(pem_text, encoding="utf-8")
        return target
    except Exception:
        return None


# Корневые сертификаты, по которым видно, что трафик проверяется «на лету»
# (SSL-инспекция антивируса или прокси), и цепочка подменяется.
#
# ВАЖНО: сертификаты этих корней нужно положить в общую склейку, иначе
# проверка TLS не пройдёт даже с сертификатом Минцифры — подменённую
# цепочку подписывает именно корень антивируса.
#
# В списке только те имена, которые НЕ встречаются у обычных публичных
# центров сертификации: Symantec, Comodo, DigiCert и прочие выдают
# сертификаты сайтам, и добавлять их в склейку нельзя — это раздувает файл
# и создаёт ложную тревогу в диагностике.
_INTERCEPTOR_MARKERS = (
    "kaspersky",
    "avast",
    "avg ",
    "eset",
    "dr.web",
    "drweb",
    "pro32",
    "bitdefender",
    "mcafee",
    "norton",
    "zscaler",
    "fortinet",
    "sophos",
    "trend micro",
    "outpost",
    "cardinal",
    "антивирус",
    "лаборатория касперского",
)

# Сколько корней SSL-инспекции максимум кладём в склейку. Одного хватает
# почти всегда, но у антивируса их может быть несколько (у Kaspersky —
# отдельные корни для «Антивируса», «Веб-Антивируса» и т.п.).
_MAX_INTERCEPTOR_ROOTS = 30


def _is_interceptor_name(subject: str) -> bool:
    """True, если Subject сертификата похож на корень антивируса/прокси."""
    lowered = subject.lower()
    return any(marker in lowered for marker in _INTERCEPTOR_MARKERS)


def _run_powershell_to_file(script: str, out_file: Path, timeout: float = 40.0) -> bool:
    """
    Запускает PowerShell-скрипт, который пишет результат в файл.

    Почему через файл, а не через stdout: вывод процесса читается либо
    напрямую, либо через канал, а канал в ограниченных окружениях бывает
    недоступен (тогда процесс молча возвращает пустоту). Файл работает
    всегда и, кроме того, не зависит от кодировки консоли: результат
    читается как UTF-8, поэтому кириллица в Subject сертификата не ломается.

    Returns:
        True, если PowerShell запустился и создал файл с данными.
    """
    try:
        out_file.unlink()
    except OSError:
        pass

    try:
        import subprocess

        completed = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                script,
            ],
            capture_output=True,
            timeout=timeout,
            check=False,
        )
    except Exception:
        logger.debug("PowerShell недоступен: %s", "запуск не удался", exc_info=True)
        return False

    if completed.returncode != 0:
        logger.debug(
            "PowerShell завершился с кодом %s: %s",
            completed.returncode,
            (completed.stderr or b"").decode("utf-8", errors="replace")[:400],
        )
    return out_file.is_file()


def _read_text_file(path: Path) -> str:
    """Читает текстовый файл в UTF-8 с запасными кодировками."""
    try:
        return read_env_text(path)
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Поиск SSL-инспекции в хранилище Windows
# ---------------------------------------------------------------------------
# Поиск идёт через PowerShell и занимает секунды (особенно на «холодной»
# машине с антивирусом). Поэтому результат кладётся в кэш, а сам поиск может
# выполняться в фоновом потоке: интерфейс не должен ждать его при запуске.
# Пустой результат тоже кэшируется — иначе каждая проверка заново запускала
# бы PowerShell.

_INTERCEPTION_CACHE: Dict[str, Any] = {"names": None, "suspects": None, "running": False}
_INTERCEPTION_LOCK = threading.Lock()


def _scan_interception_engine() -> None:
    """Один раз читает хранилище Windows и заполняет кэш SSL-инспекции."""
    try:
        names = _scan_windows_root_names()
    except Exception:
        names = []
    try:
        suspects = [subject for subject, _thumb in _interceptor_candidates(names)]
    except Exception:
        suspects = []

    with _INTERCEPTION_LOCK:
        _INTERCEPTION_CACHE["names"] = names
        _INTERCEPTION_CACHE["suspects"] = suspects
        _INTERCEPTION_CACHE["running"] = False

    if suspects:
        logger.info(
            "SSL: в корневых сертификатах Windows найдена SSL-инспекция: %s",
            "; ".join(suspects),
        )


def start_interception_scan() -> None:
    """
    Запускает поиск SSL-инспекции в фоне (один раз за запуск программы).

    Вызывается при старте интерфейса: если поиск ещё не делали, он уйдёт
    в отдельный поток и не задержит открытие окна. Результат появится
    в кэше и будет использован диагностикой.
    """
    with _INTERCEPTION_LOCK:
        if _INTERCEPTION_CACHE["names"] is not None or _INTERCEPTION_CACHE["running"]:
            return
        _INTERCEPTION_CACHE["running"] = True

    try:
        threading.Thread(
            target=_scan_interception_engine, name="ssl-interception-scan", daemon=True
        ).start()
    except Exception:
        with _INTERCEPTION_LOCK:
            _INTERCEPTION_CACHE["running"] = False


def _ensure_interception_scan() -> None:
    """Доводит поиск SSL-инспекции до конца синхронно, если он ещё не сделан."""
    with _INTERCEPTION_LOCK:
        if _INTERCEPTION_CACHE["names"] is not None:
            return
    _scan_interception_engine()


def windows_root_names() -> List[Tuple[str, str]]:
    """Корневые сертификаты Windows [(Subject, отпечаток)] с кэшированием."""
    _ensure_interception_scan()
    with _INTERCEPTION_LOCK:
        return list(_INTERCEPTION_CACHE["names"] or [])


def detect_ssl_interception() -> List[str]:
    """
    Ищет в корневых сертификатах Windows признаки SSL-инспекции.

    Возвращает список Subject-ов найденных «подменяющих» корней (например,
    «CN=Kaspersky Anti-Virus Personal Root Certificate, O=AO Kaspersky Lab»).
    Пустой список означает, что признаков нет — либо поиск ещё не завершён:
    его можно запустить заранее функцией start_interception_scan(), тогда
    проверка не задерживает интерфейс.

    Это подсказка для диагностики: антивирус может стоять и без
    SSL-инспекции, а корпоративный прокси — не иметь своего корня в системе.
    """
    if _INTERCEPTION_CACHE["suspects"] is None:
        start_interception_scan()  # не ждём: поиск идёт в фоне
    return list(_INTERCEPTION_CACHE["suspects"] or [])


def _scan_windows_root_names() -> List[Tuple[str, str]]:
    """
    Быстро читает корневые сертификаты Windows: [(Subject, отпечаток)].

    На типичной машине это ~600 сертификатов и меньше секунды работы.
    Хранилище открывается через .NET (X509Store), а не через диск Cert:\\:
    диск доступен не во всех окружениях, а .NET-класс работает всегда.

    Windows-хранилище читается только как ПОДСКАЗКА: сам Python (OpenSSL)
    его не использует, поэтому корня, установленного в системе (например,
    сертификата Минцифры или корня антивируса), программе может не хватать —
    его приходится добавлять в склейку вручную.

    Если хранилище недоступно (не Windows, запрет политики), возвращается
    пустой список: программа продолжает работать как раньше.
    """
    if os.name != "nt":  # pragma: no cover — только Windows
        return []

    # Разделитель "|" и перевод строки вместо TAB: значение Subject не может
    # содержать перевод строки, а кириллица спокойно живёт в UTF-8.
    out_file = make_work_subdir("gigachat_roots") / "windows_roots.txt"
    quoted = str(out_file).replace("'", "''")

    script = (
        "Add-Type -AssemblyName System.Security;"
        "function Get-RootItems($name) {"
        "  try {"
        "    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($name, 'LocalMachine');"
        "    $store.Open('ReadOnly');"
        "    $items = @($store.Certificates);"
        "    $store.Close();"
        "    return $items"
        "  } catch { return @() }"
        "};"
        "$all = @();"
        "$all += Get-RootItems 'Root';"
        "$all += Get-RootItems 'AuthRoot';"
        "$lines = foreach ($c in $all) {"
        "  $s = $c.Subject.Replace([char]13, ' ').Replace([char]10, ' ').Trim();"
        "  $c.Thumbprint.ToUpper() + '|' + $s"
        "};"
        f"$lines -join ([char]10) | Out-File -FilePath '{quoted}' -Encoding utf8"
    )

    if not _run_powershell_to_file(script, out_file):
        logger.debug("Список корневых сертификатов Windows получить не удалось")
        return []

    result: List[Tuple[str, str]] = []
    seen = set()

    for line in _read_text_file(out_file).splitlines():
        line = line.strip()
        if not line or "|" not in line:
            continue
        thumbprint, subject = line.split("|", 1)
        thumbprint = thumbprint.strip().upper()
        if not re.fullmatch(r"[0-9A-F]{40}", thumbprint):
            continue
        if thumbprint in seen:
            continue
        seen.add(thumbprint)
        result.append((subject.strip(), thumbprint))

    return result


def _interceptor_candidates(names: List[Tuple[str, str]]) -> List[Tuple[str, str]]:
    """
    Отбирает из списка корней те, что похожи на SSL-инспекцию.

    Возвращает [(Subject, отпечаток)] без повторов по Subject. Если
    антивирус держит несколько своих корней, попадут все (до
    _MAX_INTERCEPTOR_ROOTS) — иначе подменённая цепочка не проверится.
    """
    result: List[Tuple[str, str]] = []
    seen_subjects = set()
    for subject, thumbprint in names:
        if not _is_interceptor_name(subject):
            continue
        if subject in seen_subjects:
            continue
        seen_subjects.add(subject)
        result.append((subject, thumbprint))
        if len(result) >= _MAX_INTERCEPTOR_ROOTS:
            break
    return result


def _export_windows_roots_to_pem(
    thumbprints: List[str], out_dir: Path
) -> List[Tuple[Path, str]]:
    """
    Выгружает указанные сертификаты из хранилища Windows в PEM-файлы.

    Отпечатки передаются в PowerShell списком, поэтому на диск попадают
    ровно нужные сертификаты (обычно 1–3), а не все шестьсот.
    Проверка «issuer = subject» отсекает промежуточные сертификаты: в
    склейку нужны только корни.

    Returns:
        [(путь к PEM-файлу, Subject)] только для успешно выгруженных.
    """
    if not thumbprints:
        return []

    wanted = sorted({t.strip().upper() for t in thumbprints if t.strip()})
    if not wanted:
        return []

    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        return []

    out_file = out_dir / "_export.tsv"
    quoted = str(out_file).replace("'", "''")
    script = (
        "Add-Type -AssemblyName System.Security;"
        f"$wanted = @({','.join(repr(t) for t in wanted)});"
        "function Get-RootItems($name) {"
        "  try {"
        "    $store = New-Object System.Security.Cryptography.X509Certificates.X509Store($name, 'LocalMachine');"
        "    $store.Open('ReadOnly');"
        "    $items = @($store.Certificates);"
        "    $store.Close();"
        "    return $items"
        "  } catch { return @() }"
        "};"
        "$all = @();"
        "$all += Get-RootItems 'Root';"
        "$all += Get-RootItems 'AuthRoot';"
        "$lines = foreach ($c in $all) {"
        "  if (-not ($wanted -contains $c.Thumbprint.ToUpper())) { continue }"
        "  $b64 = [Convert]::ToBase64String($c.RawData, 'InsertLineBreaks');"
        "  $pem = \"-----BEGIN CERTIFICATE-----\" + [char]10 + $b64 + [char]10 + \"-----END CERTIFICATE-----\";"
        "  $s = $c.Subject.Replace([char]13, ' ').Replace([char]10, ' ').Trim();"
        "  $i = $c.Issuer.Replace([char]13, ' ').Replace([char]10, ' ').Trim();"
        "  $esc = $pem.Replace([char]13, [char]32).Replace([char]10, [char]126).Trim();"
        "  $c.Thumbprint.ToUpper() + '|' + $s.Replace('|', '/') + '|' + $i.Replace('|', '/') + '|' + $esc"
        "};"
        f"$lines -join ([char]10) | Out-File -FilePath '{quoted}' -Encoding utf8"
    )

    if not _run_powershell_to_file(script, out_file, timeout=45.0):
        return []

    result: List[Tuple[Path, str]] = []
    for line in _read_text_file(out_file).splitlines():
        line = line.strip()
        if not line or line.count("|") < 3:
            continue
        thumbprint, subject, issuer, pem_flat = line.split("|", 3)
        thumbprint = thumbprint.strip().upper()
        if not re.fullmatch(r"[0-9A-F]{40}", thumbprint):
            continue
        if subject.strip() != issuer.strip():
            continue  # это не корень (сам себе не подписан) — в склейку не берём

        # В PowerShell перевод строки заменён на "~" (символ, которого в PEM
        # не бывает), поэтому обратная замена однозначна.
        pem = pem_flat.strip().replace("~", "\n").strip() + "\n"
        if "BEGIN CERTIFICATE" not in pem:
            continue

        # Сверяем отпечаток: файл должен быть тем самым сертификатом.
        try:
            der = ssl.PEM_cert_to_DER_cert(pem)
        except Exception:
            continue
        if hashlib.sha1(der).hexdigest().upper() != thumbprint:
            continue

        pem_path = out_dir / f"{thumbprint.lower()}.pem"
        try:
            pem_path.write_text(pem, encoding="utf-8")
        except Exception:
            continue
        result.append((pem_path, subject.strip()))

    if not result:
        logger.debug(
            "Ни один корень SSL-инспекции не выгружен из хранилища Windows "
            "(отпечатков запрошено: %s)",
            len(wanted),
        )
    return result


def _read_cached_interceptor_pems(out_dir: Path) -> List[Tuple[Path, str]]:
    """
    Читает уже выгруженные корни SSL-инспекции (без обращения к Windows).

    Имена файлов — отпечатки сертификатов этого и хватает, чтобы не
    запускать PowerShell заново: содержимое проверяется при первом создании
    файла, а папка живёт во временном каталоге пользователя.
    """
    result: List[Tuple[Path, str]] = []
    try:
        entries = sorted(out_dir.glob("*.pem"))
    except OSError:
        return result

    for path in entries:
        try:
            pem = path.read_text(encoding="utf-8", errors="ignore")
            der = ssl.PEM_cert_to_DER_cert(pem)
        except Exception:
            continue
        if hashlib.sha1(der).hexdigest().upper() != path.stem.upper():
            continue  # имя не совпало с содержимым — файл не наш
        result.append((path, _cert_common_name(der) or path.stem))
    return result


def _interceptor_pem_files(exclude: Optional[Path] = None) -> List[Tuple[Path, str]]:
    """
    PEM-файлы корней SSL-инспекции из хранилища Windows (для склейки).

    Именно эти сертификаты подписывают подменённую цепочку, поэтому без них
    проверка TLS не пройдёт даже с сертификатом Минцифры.

    Сначала проверяются уже выгруженные файлы (быстро, без запуска
    PowerShell), и только если их нет — читается хранилище Windows.
    """
    out_dir = make_work_subdir("gigachat_roots")

    cached = _read_cached_interceptor_pems(out_dir)
    if cached:
        return [(path, subject) for path, subject in cached if path != exclude]

    names = windows_root_names()
    candidates = _interceptor_candidates(names)
    if not candidates:
        return []

    exported = _export_windows_roots_to_pem([t for _s, t in candidates], out_dir)

    return [(path, subject) for path, subject in exported if path != exclude]


def _locate_av_root_pem(exclude: Optional[Path] = None) -> Optional[Tuple[Path, str]]:
    """
    Возвращает первый корень SSL-инспекции (путь к PEM, Subject) или None.

    Совместимая обёртка над _interceptor_pem_files().
    """
    found = _interceptor_pem_files(exclude)
    return found[0] if found else None


def build_ca_bundle(logger_: Optional[Any] = None) -> Optional[Path]:
    """
    Собирает ОДИН файл-склейку из всех нужных корневых сертификатов.

    Зачем склейка, а не один файл. Параметр verify у httpx принимает либо
    набор системных корней (certifi), либо ОДИН файл. Как только программа
    подставляет сертификат Минцифры, certifi перестаёт использоваться — и
    любая цепочка, которой нужен обычный корень (или корень антивируса при
    SSL-инспекции), начинает падать. Поэтому в склейку попадают:

      1. корень Минцифры (без него GigaChat не проверяется вообще);
      2. дополнительно — все *.cer/*.crt/*.pem из config\\ (корень
         антивируса, корпоративный корень и т.п.);
      3. штатный cacert.pem от certifi — чтобы не потерять обычные корни;
      4. корень SSL-инспекции из хранилища Windows, если он там есть, —
         именно он подписывает подменённую цепочку.

    Сертификаты сравниваются по SHA-1 отпечатку, дубликаты не попадают.
    Файл перезаписывается только при изменении содержимого: лишних
    обращений к диску нет.

    Returns:
        Путь к склейке или None, если ни одного сертификата не нашлось.
    """
    log = logger_ or logger

    import tempfile

    config_dir = PROJECT_ROOT / "config"
    if IS_FROZEN:
        try:
            config_dir = Path(sys.executable).resolve().parent / "config"
        except Exception:
            pass
    target = config_dir / MERGED_CA_BUNDLE_FILENAME

    # Куда писать, если config\ недоступна (например, программа лежит в
    # защищённой папке): временная папка. Работает, хотя и менее заметно.
    if not _is_path_inside(target, PROJECT_ROOT) and not IS_FROZEN:
        fallback_dir = Path(os.environ.get("TEMP") or os.getcwd())
        target = fallback_dir / MERGED_CA_BUNDLE_FILENAME

    candidates: List[Path] = []

    def add(path: Optional[Path]) -> None:
        if path is None:
            return
        try:
            if path.is_file() and path not in candidates:
                candidates.append(path)
        except OSError:
            return

    # GIGACHAT_CA_BUNDLE_FILE мог быть задан явно — в склейке он первый
    explicit = (os.environ.get("GIGACHAT_CA_BUNDLE_FILE") or "").strip()
    if explicit:
        add(Path(explicit))

    add(find_ca_bundle())
    for extra in _extra_ca_files():
        add(extra)
    add(_certifi_bundle())

    av_roots = _interceptor_pem_files(exclude=target)
    for av_path, av_subject in av_roots:
        add(av_path)

    if not candidates:
        return None

    # Отпечаток источников: пути, размеры и время изменения. По нему видно,
    # изменился ли хоть один сертификат с прошлого запуска. Это позволяет
    # НЕ пересобирать склейку на каждом старте программы: сборка заново
    # читает ~120 сертификатов и заметно тормозит запуск.
    fingerprint_parts: List[str] = []
    for path in candidates:
        try:
            stat = path.stat()
            fingerprint_parts.append(f"{path}|{stat.st_size}|{int(stat.st_mtime)}")
        except OSError:
            continue
    sources_fingerprint = hashlib.sha1(
        "\n".join(fingerprint_parts).encode("utf-8", errors="replace")
    ).hexdigest()

    def is_current(candidate_file: Path) -> bool:
        """True, если готовая склейка собрана из тех же источников."""
        try:
            if not candidate_file.is_file():
                return False
            with candidate_file.open("r", encoding="utf-8", errors="ignore") as handle:
                for _ in range(6):
                    line = handle.readline()
                    if not line:
                        break
                    if line.startswith("# sources="):
                        return line.strip().endswith(sources_fingerprint)
        except Exception:
            return False
        return False

    for known_path in (target, get_work_dir() / MERGED_CA_BUNDLE_FILENAME):
        if is_current(known_path):
            return known_path

    # Собираем блоки сертификатов по одному, сохраняя порядок источников:
    # Минцифры → свои сертификаты из config\ → корень антивируса → certifi.
    # Дубликаты отсекаются по SHA-1 отпечатку.
    merged_blocks: List[str] = []
    merged_thumbs = set()
    for path in candidates:
        try:
            raw = path.read_bytes()
        except Exception:
            continue

        blocks = re.findall(
            r"-----BEGIN CERTIFICATE-----.*?-----END CERTIFICATE-----",
            raw.decode("utf-8", errors="ignore"),
            re.DOTALL,
        )
        if not blocks and raw[:1] == b"\x30":
            try:
                blocks = [ssl.DER_cert_to_PEM_cert(raw).strip()]
            except Exception:
                blocks = []

        for block in blocks:
            try:
                der = ssl.PEM_cert_to_DER_cert(block)
            except Exception:
                continue
            thumb = hashlib.sha1(der).hexdigest()
            if thumb in merged_thumbs:
                continue
            merged_thumbs.add(thumb)
            merged_blocks.append(block.strip())

    if not merged_blocks:
        return None

    header = (
        "# Автоматически собранный набор корневых сертификатов «Генератора КП».\n"
        "# Файл создаётся программой при запуске, править его вручную не нужно.\n"
        f"# Корней в наборе: {len(merged_blocks)}. Источники: config\\, certifi, "
        "хранилище Windows.\n"
        f"# sources={sources_fingerprint}\n"
    )
    content = header + "\n".join(merged_blocks) + "\n"

    def try_write(path: Path) -> Optional[Path]:
        """
        Пытается записать склейку по указанному пути.

        Файл не переписывается, если содержимое не изменилось: при каждом
        запуске программы лишних обращений к диску нет.
        """
        try:
            if path.is_file() and path.read_text(encoding="utf-8", errors="ignore") == content:
                return path
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(content, encoding="utf-8")
            return path
        except Exception:
            return None

    # Порядок попыток: сначала config\ рядом с программой (файл виден
    # пользователю), затем рабочая папка. Папка программы может быть
    # защищена от записи (Program Files, права администратора, антивирус) —
    # тогда программа всё равно должна получить рабочий набор сертификатов,
    # а не остаться без проверки TLS.
    written = try_write(target)
    if written is None:
        fallback = get_work_dir() / MERGED_CA_BUNDLE_FILENAME
        written = try_write(fallback)
        if written is not None:
            log.info(
                "SSL: папка config\\ недоступна для записи (%s) — набор "
                "сертификатов собран в рабочей папке: %s",
                target,
                written,
            )

    if written is None:
        log.warning("SSL: не удалось собрать объединённый набор сертификатов")
        return None

    if written == target:
        log.info(
            "SSL: сертификаты объединены в один файл (%s корней): %s",
            len(merged_blocks),
            written,
        )
    return written


def setup_ca_bundle(logger_: Optional[Any] = None) -> Optional[str]:
    """
    Подставляет сертификаты в GIGACHAT_CA_BUNDLE_FILE, если они найдены.

    Логика:
      * значение, заданное в .env, НЕ переопределяется — путь пользователя
        всегда в приоритете (но и он проходит через склейку, чтобы не
        потерять остальные корни);
      * иначе собирается склейка (см. build_ca_bundle) и её путь попадает
        в GIGACHAT_CA_BUNDLE_FILE.

    Returns:
        Путь к используемому набору сертификатов или None.
    """
    log = logger_ or logger

    explicit = (os.environ.get("GIGACHAT_CA_BUNDLE_FILE") or "").strip()
    if explicit and explicit != "none" and os.path.isfile(explicit):
        log.info("CA-сертификат задан явно в .env: %s", explicit)

    problem = get_ca_bundle_problem()
    if problem:
        log.warning("SSL: %s", problem)

    merged = build_ca_bundle(log)
    if merged is not None:
        os.environ["GIGACHAT_CA_BUNDLE_FILE"] = str(merged)
        return str(merged)

    if explicit:
        log.warning("GIGACHAT_CA_BUNDLE_FILE указывает на несуществующий файл: %s", explicit)
        return None

    search_paths = _ca_bundle_search_paths()
    log.info(
        "CA-сертификат Минцифры не найден — искали в %s местах (первое: %s)",
        len(search_paths),
        search_paths[0],
    )
    return None


def get_ca_bundle_status() -> Dict[str, Any]:
    """
    Полный статус сертификатов — для диагностики в интерфейсе и логе.

    Returns:
        found:          путь к найденному сертификату Минцифры или None;
        expected:       путь, где сертификат ищут в первую очередь;
        search_paths:   все места поиска (в порядке приоритета);
        searched_count: сколько мест проверено;
        explicit:       путь, заданный вручную в GIGACHAT_CA_BUNDLE_FILE;
        explicit_valid: существует ли файл, заданный вручную;
        problem:        текст проблемы (пустой/битый файл), если она есть;
        merged_bundle:  путь к собранной склейке сертификатов (или None);
        merged_count:   сколько корней в склейке;
        interception:   Subject-ы корней антивируса из хранилища Windows;
        store_scan:     удалось ли прочитать хранилище Windows
                        («found» / «empty» / «skipped»);
        verify_ssl:     включена ли проверка сертификата;
        url:            официальный адрес загрузки сертификата.
    """
    found = find_ca_bundle()
    explicit = (os.environ.get("GIGACHAT_CA_BUNDLE_FILE") or "").strip()
    explicit_valid = bool(explicit) and os.path.isfile(explicit)

    search_paths = [str(path) for path in _ca_bundle_search_paths()]

    # Склейку собираем только при необходимости: чтение хранилища Windows
    # идёт через PowerShell и занимает время — в диагностике это заметно.
    merged = explicit if explicit_valid else None
    merged_count = 0
    if merged:
        try:
            merged_count = len(
                re.findall(
                    r"-----BEGIN CERTIFICATE-----",
                    read_env_text(Path(merged)),
                )
            )
        except Exception:
            merged_count = 0

    # Признаки SSL-инспекции берём из кэша. Если фоновый поиск ещё идёт,
    # кэш пуст — это не ошибка: диагностика покажет результат следующего
    # открытия. Запускать PowerShell прямо здесь нельзя: окно диагностики
    # должно открываться мгновенно.
    with _INTERCEPTION_LOCK:
        cached_names = _INTERCEPTION_CACHE["names"]
        interception = list(_INTERCEPTION_CACHE["suspects"] or [])

    if os.name != "nt":
        store_scan = "skipped"
    elif isinstance(cached_names, list) and cached_names:
        store_scan = "found"
    else:
        store_scan = "empty"

    return {
        "found": str(found) if found else None,
        "expected": str(_ca_bundle_search_paths()[0]),
        "search_paths": search_paths,
        "searched_count": len(search_paths),
        "explicit": explicit or None,
        "explicit_valid": explicit_valid,
        "problem": get_ca_bundle_problem(),
        "merged_bundle": merged,
        "merged_count": merged_count,
        "interception": interception,
        "store_scan": store_scan,
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

    if status.get("interception"):
        log.error(
            "SSL: в корневых сертификатах Windows найден корень SSL-инспекции "
            "(антивирус или прокси подменяет сертификат):"
        )
        for subject in status["interception"]:
            log.error("SSL:   - %s", subject)
        log.error(
            "SSL: этот корень добавлен в общий набор сертификатов программы. "
            "Если ошибка осталась — добавьте программу в исключения "
            "SSL-инспекции антивируса."
        )
    elif status.get("store_scan") == "empty":
        log.error(
            "SSL: список корневых сертификатов Windows прочитать не удалось — "
            "проверить SSL-инспекцию автоматически нельзя"
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

    Текст собирается по фактическому состоянию: если сертификат уже найден,
    не предлагаем его «скачать и положить», а объясняем следующую причину
    (SSL-инспекция). Если SSL-инспекция обнаружена, называем виновника и
    говорим, что программа уже добавила его корень в набор сертификатов.
    """
    status = {}
    try:
        status = get_ca_bundle_status()
    except Exception:
        pass

    ca_path = PROJECT_ROOT / "config" / CA_BUNDLE_FILENAME
    found = status.get("found")
    interception = status.get("interception") or []
    merged = status.get("merged_bundle")

    lines: List[str] = ["⚠️ SSL-ошибка при обращении к GigaChat.", ""]

    if found:
        lines += [
            "Сертификат Минцифры найден:",
            f"   {found}",
            "",
        ]
    else:
        lines += [
            "Корневой сертификат Минцифры НЕ найден.",
            "1. Скачайте сертификат:",
            f"   {CA_BUNDLE_URL}",
            "2. Положите его в папку config рядом с программой:",
            f"   {ca_path}",
            "   (или просто рядом с программой — программа ищет и там)",
            "3. Либо пропишите свой путь в .env:",
            "   GIGACHAT_CA_BUNDLE_FILE=C:\\certs\\russian_trusted_root_ca.cer",
            "",
        ]

    if interception:
        lines += [
            "🔎 Обнаружена SSL-инспекция — трафик проверяется на лету:",
        ]
        for subject in interception:
            lines.append(f"   - {subject.split(',')[0].replace('CN=', '').strip()}")
        lines += [
            "Корень этой программы уже добавлен в общий набор сертификатов,",
            "поэтому обычно ничего делать не нужно — перезапустите программу.",
            "",
            "Если ошибка осталась, исключите программу из SSL-инспекции:",
            "   Kaspersky: Настройки → Сеть → Проверка защищённых соединений",
            "              → Исключения → добавить Генератор_КП_GigaChat.exe",
            "   Pro32:     Настройки → Защита → Исключения / SSL-инспекция",
            "",
        ]
    else:
        lines += [
            "Второй возможный виновник — антивирус или корпоративный прокси,",
            "который подменяет сертификат своим (SSL-инспекция).",
            "Добавьте программу в его исключения:",
            "   Kaspersky: Настройки → Сеть → Проверка защищённых соединений",
            "              → Исключения",
            "   Pro32:     Настройки → Защита → Исключения / SSL-инспекция",
            "",
        ]

    if merged:
        lines += [
            "Программа проверяет TLS по своему набору сертификатов:",
            f"   {merged}",
            f"   корней в наборе: {status.get('merged_count') or '?'}",
            "",
        ]

    lines += [
        "Отключать GIGACHAT_VERIFY_SSL_CERTS не нужно: это убирает защиту",
        "от подмены трафика, а проблему не решает.",
    ]
    return "\n".join(lines)


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

# Таймаут одного запроса к GigaChat, секунды.
#
# 60 с — это «потолок ожидания», а не обычное время ответа (нормальный ответ
# приходит за 1–3 с). Запас нужен для медленных сетей: при 30 с запрос мог
# обрываться ровно на середине генерации. Переопределяется GIGACHAT_TIMEOUT
# в .env.
DEFAULT_GIGACHAT_TIMEOUT = 60.0

# Адрес API GigaChat.
#
# ВАЖНО, это главная причина ошибки [WinError 10060].
# Начиная с версии 0.2.0 библиотека gigachat по умолчанию ходит на
# https://api.giga.chat/v1 (PR #119 проекта ai-forever/gigachat). Этот адрес
# доступен НЕ из всех сетей: TCP-соединение к нему просто не устанавливается
# (пакеты отбрасываются), и запрос падает с ConnectTimeout [WinError 10060],
# хотя авторизация при этом проходит — она идёт на другой хост и работает.
#
# Классический адрес Сбера работает стабильно, поэтому задан явно и по
# умолчанию. Переопределяется переменной GIGACHAT_BASE_URL в .env — тогда
# можно увести трафик на прокси или на новый адрес, не трогая код.
DEFAULT_GIGACHAT_BASE_URL = "https://gigachat.devices.sberbank.ru/api/v1"

# Сколько раз повторять запрос при сетевом таймауте и пауза между попытками.
# Первая попытка + один повтор: соединение иногда не устанавливается с
# первого раза (файрвол, антивирус, нестабильный канал), а второй заход
# проходит нормально.
_GIGACHAT_ATTEMPTS = 2
_GIGACHAT_RETRY_DELAY = 1.5

# ============================================
# ТИП ОПЛАТЫ (специфика лотов для автовозов)
# ============================================
# Ключ -> (надпись в интерфейсе, формулировка для текста КП).
# Ключи попадают в settings.json, поэтому менять их без миграции не стоит.
PAYMENT_TYPES = {
    "beznal_nds": ("Безнал с НДС", "Оплата по безналу (с НДС)."),
    "beznal": ("Безнал без НДС", "Оплата по безналу (без НДС)."),
    "cash": ("Наличные", "Оплата наличными (без НДС)."),
    "discuss": ("Обсуждается", "Оплата обсуждается — безнал с НДС или наличные."),
}
PAYMENT_ORDER = list(PAYMENT_TYPES.keys())
DEFAULT_PAYMENT_KEY = "beznal_nds"

# Надписи для выпадающего списка (по ним же ищем ключ при выборе)
PAYMENT_LABELS = [label for label, _ in PAYMENT_TYPES.values()]


def payment_key_from_label(label: str) -> str:
    """Превращает надпись из списка в ключ. Неизвестная надпись → оплата по умолчанию."""
    for key, (shown, _) in PAYMENT_TYPES.items():
        if shown == label:
            return key
    return DEFAULT_PAYMENT_KEY


def payment_line(key: str) -> str:
    """Готовая строка про оплату для текста КП."""
    return PAYMENT_TYPES.get(key, PAYMENT_TYPES[DEFAULT_PAYMENT_KEY])[1]


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

# Имена исключений, означающие «соединение не установилось / оборвалось».
# Такие запросы имеет смысл повторить: сервис тут ни при чём, дело в сети.
_CONNECT_ERROR_NAMES = {
    "ConnectTimeout",
    "ConnectError",
    "ConnectionError",
    "ConnectionResetError",
    "ReadTimeout",
    "TimeoutException",
    "Timeout",
}

# Текстовые маркеры того же самого — на случай, если ошибка пришла обёрнутой
# в чужое исключение и по имени типа её не узнать.
_CONNECT_ERROR_MARKERS = (
    "winerror 10060",
    "connecttimeout",
    "connect timeout",
    "connection timed out",
    "connection aborted",
)

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
    """
    Возвращает Authorization Key из окружения (после загрузки .env).

    Значение проходит через sanitize_key(): невидимые символы и случайные
    пробелы, попавшие в .env при копировании, больше не превращают рабочий
    ключ в «неверный».
    """
    return sanitize_key(os.environ.get("GIGACHAT_CREDENTIALS"))


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


def _resolve_base_url() -> str:
    """
    Возвращает адрес API GigaChat.

    Приоритет: GIGACHAT_BASE_URL из .env → DEFAULT_GIGACHAT_BASE_URL.
    Пустое значение в .env считается «не задано»: раньше пустая строка
    означала «взять адрес библиотеки по умолчанию», а он (api.giga.chat)
    доступен не из всех сетей — из-за этого запросы и падали.
    """
    configured = (os.environ.get("GIGACHAT_BASE_URL") or "").strip()
    return configured or DEFAULT_GIGACHAT_BASE_URL


def _current_config_key() -> tuple:
    """Отпечаток конфигурации: при её смене клиент пересоздаётся."""
    return (
        get_gigachat_credentials(),
        os.environ.get("GIGACHAT_SCOPE", DEFAULT_GIGACHAT_SCOPE),
        os.environ.get("GIGACHAT_MODEL", DEFAULT_GIGACHAT_MODEL),
        os.environ.get("GIGACHAT_VERIFY_SSL_CERTS", "true"),
        os.environ.get("GIGACHAT_CA_BUNDLE_FILE", ""),
        # Адрес входит в отпечаток: сменили GIGACHAT_BASE_URL — клиент
        # пересоздастся сам, старый останется ходить на прежний хост.
        _resolve_base_url(),
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


# Невидимые символы, которые попадают в ключ при копировании из браузера и
# из-за которых сервер отвечает «неверный ключ», хотя внешне строка верная.
_INVISIBLE_CHARS = (
    "\u200b",  # нулевой пробел
    "\u200c",  # нулевой неприсоединяемый
    "\u200d",  # нулевой присоединяемый
    "\u2060",  # word joiner
    "\ufeff",  # BOM
    "\u00a0",  # неразрывный пробел
)


def sanitize_key(value: Optional[str]) -> str:
    """
    Приводит введённый ключ к виду, который понимает GigaChat.

    Что убирается и почему:
      * пробелы, табы и переводы строк — при вставке из письма или чата
        ключ легко приезжает «с хвостом», и сервер его не принимает;
      * кавычки по краям — многие копируют ключ вместе с ними;
      * невидимые символы (нулевой пробел, BOM, неразрывный пробел) —
        визуально ключ верный, а по байтам нет. Раньше такой ключ
        сохранялся как есть, и программа сообщала «неверный ключ».

    Само значение не изменяется: допустимые символы ключа (base64 и знаки
    «-», «_») остаются нетронутыми.
    """
    if value is None:
        return ""

    text = str(value)
    for ch in _INVISIBLE_CHARS:
        text = text.replace(ch, "")
    text = "".join(ch for ch in text if ch.isprintable())

    return text.strip().strip("'\"").strip()


def save_gigachat_key(key: str) -> Path:
    """
    Сохраняет ключ GigaChat в .env, который программа реально читает.

    Пишет в ТОТ ЖЕ файл, который ищет find_env_file()/get_env_write_path():
    для собранного EXE — рядом с .exe, для исходников — в корне проекта.
    Так запись и чтение никогда не расходятся (раньше ключ мог сохраняться
    в одном месте, а читаться из другого — и ввод «не работал»).

    Остальные настройки .env не затираются: строка GIGACHAT_CREDENTIALS
    заменяется или добавляется, недостающие ключи создаются. После записи
    переменная окружения в текущем процессе обновляется, а кэш клиента
    сбрасывается — новый ключ начинает действовать без перезапуска.

    Args:
        key: значение Authorization Key (сам ключ в лог не попадает).

    Returns:
        Path к записанному файлу .env.

    Raises:
        ValueError: если после очистки от лишних символов ключ пуст.
    """
    key = sanitize_key(key)
    if not key:
        raise ValueError("Ключ пуст")

    env_path = get_env_write_path()
    env_path.parent.mkdir(parents=True, exist_ok=True)

    lines: List[str] = []
    if env_path.exists():
        try:
            # read_env_text устойчива к BOM и кодировкам cp1251/cp866
            lines = [
                line.rstrip("\n")
                for line in read_env_text(env_path).splitlines()
                if not line.startswith("GIGACHAT_CREDENTIALS")
            ]
        except Exception:
            lines = []

    if not lines:
        lines = [
            "GIGACHAT_SCOPE=GIGACHAT_API_PERS",
            "GIGACHAT_MODEL=GigaChat",
            f"GIGACHAT_TIMEOUT={int(DEFAULT_GIGACHAT_TIMEOUT)}",
            "GIGACHAT_VERIFY_SSL_CERTS=true",
            "GIGACHAT_CA_BUNDLE_FILE=",
            f"GIGACHAT_BASE_URL={DEFAULT_GIGACHAT_BASE_URL}",
        ]
    else:
        # Обе настройки дописываем независимо друг от друга: старый .env
        # может не содержать ни одной, ни другой. Без GIGACHAT_BASE_URL
        # библиотека уходит на недоступный api.giga.chat.
        if not any(line.startswith("GIGACHAT_CA_BUNDLE_FILE") for line in lines):
            lines.append("GIGACHAT_CA_BUNDLE_FILE=")
        if not any(line.startswith("GIGACHAT_BASE_URL") for line in lines):
            lines.append(f"GIGACHAT_BASE_URL={DEFAULT_GIGACHAT_BASE_URL}")

    lines.append(f"GIGACHAT_CREDENTIALS={key}")

    # encoding="utf-8" без BOM: так файл читают и dotenv, и сам генератор
    env_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    # Подхватываем новый ключ без перезапуска программы
    os.environ["GIGACHAT_CREDENTIALS"] = key
    reset_gigachat_client()

    # ВАЖНО: сам ключ в лог не пишем
    logger.info("Ключ сохранён в %s", env_path)
    return env_path


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

    base_url = _resolve_base_url()

    client_kwargs: Dict[str, Any] = {
        "credentials": credentials,
        "scope": scope,
        "model": model,
        "timeout": timeout,
        "verify_ssl_certs": verify_ssl,
        "ca_bundle_file": ca_bundle,
        # Адрес задаём ВСЕГДА и явно. Полагаться на значение по умолчанию
        # нельзя: с gigachat 0.2.0 это api.giga.chat, недоступный из части
        # сетей (см. DEFAULT_GIGACHAT_BASE_URL).
        "base_url": base_url,
    }

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
            "base_url=%s | timeout=%sс | verify_ssl=%s | ca_bundle=%s",
            mask_credentials(credentials),
            scope,
            model,
            base_url,
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
    Потоковая генерация (stream=True): текст приходит по мере готовности.

    ВНИМАНИЕ: по умолчанию НЕ используется, и это осознанно.
    В сетях с SSL-инспекцией (Kaspersky, Pro32 и подобные) подмена
    сертификата рвёт длинную HTTPS-сессию: соединение открывается, но поток
    чанков не доходит, и запрос умирает по ConnectTimeout. Именно поэтому
    основной путь — _generate_gigachat_sync(), а эта функция вызывается
    только при явном use_stream=True (например, в сети без инспекции).

    Хвост «stream» оставлен как возможность, а не как рабочий режим.
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
    """Обычная (непотоковая) генерация — основной режим работы."""
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


def is_connect_timeout_error(error: Any) -> bool:
    """
    Ошибка означает «соединение не установилось / оборвалось по сети».

    Такие сбои имеет смысл повторять: сервис тут ни при чём, дело в канале,
    файрволе или антивирусе, и вторая попытка часто проходит.

    SSL-ошибки сюда НЕ попадают: там проблема в сертификате, и повтор
    ничего не изменит — только зря прождём таймаут.
    """
    if is_ssl_error(error):
        return False

    if type(error).__name__ in _CONNECT_ERROR_NAMES:
        return True

    text = scrub_text(str(error)).lower()
    return any(marker in text for marker in _CONNECT_ERROR_MARKERS)


def _generate_gigachat_sync_with_retry(
    client: Any,
    system_prompt: str,
    prompt: str,
    temperature: float,
    max_tokens: int,
    attempts: int = _GIGACHAT_ATTEMPTS,
) -> str:
    """
    Синхронная генерация с одной повторной попыткой при сетевом таймауте.

    Повторяем ТОЛЬКО сетевые сбои (ConnectTimeout / WinError 10060 и родню).
    Ошибку авторизации, SSL или ответ 4xx повторять бессмысленно — она
    повторится ровно так же, а пользователь прождёт лишний таймаут.

    Если повтор не помог, наружу уходит последняя ошибка: вызывающий код
    логирует её и уходит в шаблонный текст (fallback сохранён).
    """
    last_error: Optional[BaseException] = None

    for attempt in range(1, max(1, attempts) + 1):
        try:
            return _generate_gigachat_sync(
                client, system_prompt, prompt, temperature, max_tokens
            )
        except Exception as e:
            last_error = e

            if attempt >= attempts or not is_connect_timeout_error(e):
                raise

            logger.warning(
                "GigaChat: сетевой сбой (%s), попытка %d из %d — "
                "повторяю запрос через %.1f сек",
                scrub_text(str(e))[:200],
                attempt,
                attempts,
                _GIGACHAT_RETRY_DELAY,
            )
            time.sleep(_GIGACHAT_RETRY_DELAY)

    # Сюда попасть нельзя: цикл либо вернёт текст, либо выбросит исключение.
    raise last_error if last_error is not None else RuntimeError("GigaChat: нет попыток")


def generate_gigachat(
    prompt: str,
    system_prompt: Optional[str] = None,
    temperature: float = 0.85,
    max_tokens: int = 280,
    use_stream: bool = False,
    timeout: Optional[float] = None,
) -> Optional[str]:
    """
    Генерирует текст через GigaChat.

    Args:
        prompt:        пользовательский запрос
        system_prompt: системная роль (по умолчанию — копирайтер по автоперевозкам)
        temperature:   креативность (0.0–2.0)
        max_tokens:    ограничение длины ответа (меньше = быстрее)
        use_stream:    ВЫКЛЮЧЕНО по умолчанию. Потоковая генерация в сетях
                       с SSL-инспекцией рвётся по таймауту, поэтому рабочий
                       режим — обычный запрос. Включать только осознанно.
        timeout:       таймаут запроса в секундах; None — взять GIGACHAT_TIMEOUT
                       из .env (по умолчанию 60 с). Запрос ВСЕГДА ограничен
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
            # Явный запрос стриминга. Оставлен для сетей без SSL-инспекции,
            # но по умолчанию этот путь не задействован.
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
                text = _generate_gigachat_sync_with_retry(
                    client, system_prompt, str(prompt), temperature, max_tokens
                )
        else:
            text = _generate_gigachat_sync_with_retry(
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
    "Ты — логист по автовозам. Пиши короткие коммерческие предложения для "
    "перевозчиков-автовозов: опиши лот машин, который нужно перевезти автовозом "
    "из точки А в точку Б. Пиши на русском языке, энергично, как в чате перевозчиков, "
    "без ошибок и лишней воды. "
    "ЗАПРЕЩЕНО упоминать: цену, количество машин, названия марок и моделей машин. "
    "ОБЯЗАТЕЛЬНО: слово «автовоз» рядом с упоминанием лота, маршрут со всеми городами, "
    "тип оплаты и призыв писать в личку. "
    "Не выдумывай названия компаний и сроки. Не добавляй подписи и обращения."
)


# ============================================
# ЕДИНАЯ ТОЧКА ВХОДА
# ============================================


def get_available_provider() -> str:
    """
    Определяет доступный провайдер.

    Returns:
        "gigachat" | "none"
    """
    if is_gigachat_configured():
        return "gigachat"
    return "none"


def generate_text(
    prompt: str,
    provider: str = "gigachat",
    system_prompt: Optional[str] = None,
    temperature: float = 0.85,
    max_tokens: int = 280,
) -> Optional[str]:
    """
    Единая точка генерации текста.

    provider="gigachat" — GigaChat
    provider="auto"     — то же самое: других бэкендов в проекте нет

    Returns:
        Текст или None, если GigaChat не сработал.
    """
    if provider in ("auto", "gigachat"):
        return generate_gigachat(
            prompt=prompt,
            system_prompt=system_prompt,
            temperature=temperature,
            max_tokens=max_tokens,
        )

    logger.warning("generate_text: неизвестный провайдер (provider=%s)", provider)
    return None


def get_status_info() -> Dict[str, Any]:
    """Краткая сводка о состоянии провайдера — для логов и UI."""
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

# ВАЖНО: при импорте модуля НИЧЕГО не собирается и не сканируется.
# Раньше здесь вызывался setup_ca_bundle(), который читал хранилище
# сертификатов Windows через PowerShell — на машине с антивирусом это
# занимало десятки секунд, и всё это время окно программы не открывалось
# (импорт идёт в том же потоке, что и построение интерфейса).
#
# Теперь набор сертификатов собирается лениво: при создании клиента
# GigaChat (_get_gigachat_client), то есть в фоновом потоке при прогреве
# или генерации. Если сертификат Минцифры уже есть, при старте он просто
# попадает в GIGACHAT_CA_BUNDLE_FILE — это мгновенная проверка файла.

try:
    _startup_ca = find_ca_bundle()
    if _startup_ca is not None:
        os.environ.setdefault("GIGACHAT_CA_BUNDLE_FILE", str(_startup_ca))
except Exception:  # pragma: no cover — поиск не должен мешать запуску
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
