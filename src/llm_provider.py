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

import hashlib
import os
import sys
import re
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


def _copy_windows_cert_to_pem(cert_id: str, pem_text: str) -> Optional[Path]:
    """
    Кладёт PEM-текст сертификата в стабильный файл во временной папке.

    Имя файла включает отпечаток, поэтому файл создаётся один раз и потом
    просто переиспользуется. Возвращает путь или None.
    """
    import tempfile

    safe = re.sub(r"[^0-9A-Fa-f]", "", cert_id)[:40] or hashlib.sha1(
        pem_text.encode("utf-8")
    ).hexdigest()[:16]
    target = Path(tempfile.gettempdir()) / f"gigachat_extra_root_{safe.upper()}.pem"
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

    import tempfile

    # Разделитель "|" и перевод строки вместо TAB: значение Subject не может
    # содержать перевод строки, а кириллица спокойно живёт в UTF-8.
    out_file = Path(tempfile.gettempdir()) / "gigachat_windows_roots.txt"
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


def detect_ssl_interception() -> List[str]:
    """
    Ищет в корневых сертификатах Windows признаки SSL-инспекции.

    Возвращает список Subject-ов найденных «подменяющих» корней (например,
    «CN=Kaspersky Anti-Virus Personal Root Certificate, O=AO Kaspersky Lab»).
    Пустой список означает, что признаков нет.

    Это подсказка для диагностики: антивирус может стоять и без
    SSL-инспекции, а корпоративный прокси — не иметь своего корня в системе.
    """
    return [subject for subject, _thumb in _interceptor_candidates(_scan_windows_root_names())]


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
        "  $der = $c.RawData;"
        "  $b64 = [Convert]::ToBase64String($der, 'InsertLineBreaks');"
        "  $pem = \"-----BEGIN CERTIFICATE-----\" + [Environment]::NewLine + $b64 + "
        "[Environment]::NewLine + \"-----END CERTIFICATE-----\";"
        "  $s = $c.Subject.Replace([char]13, ' ').Replace([char]10, ' ').Trim();"
        "  $i = $c.Issuer.Replace([char]13, ' ').Replace([char]10, ' ').Trim();"
        "  $esc = $pem.Replace([char]13, '').Replace([char]10, '\\n');"
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

        pem = pem_flat.replace("\\n", "\n").strip() + "\n"
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


def _interceptor_pem_files(exclude: Optional[Path] = None) -> List[Tuple[Path, str]]:
    """
    PEM-файлы корней SSL-инспекции из хранилища Windows (для склейки).

    Именно эти сертификаты подписывают подменённую цепочку, поэтому без них
    проверка TLS не пройдёт даже с сертификатом Минцифры.
    """
    names = _scan_windows_root_names()
    candidates = _interceptor_candidates(names)
    if not candidates:
        return []

    import tempfile

    out_dir = Path(tempfile.gettempdir()) / "gigachat_roots"
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
        log.info("SSL: корень SSL-инспекции (антивирус/прокси) добавлен в склейку: %s", av_subject)
        add(av_path)

    if not candidates:
        return None

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
    )
    content = header + "\n".join(merged_blocks) + "\n"

    try:
        if target.is_file() and target.read_text(encoding="utf-8", errors="ignore") == content:
            return target  # ничего не изменилось — файл не трогаем
    except Exception:
        pass

    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(content, encoding="utf-8")
        log.info(
            "SSL: сертификаты объединены в один файл (%s корней): %s",
            len(merged_blocks),
            target,
        )
        return target
    except Exception as exc:
        log.warning("SSL: не удалось записать объединённый сертификат %s: %s", target, exc)

    # Последняя попытка — во временной папке
    try:
        import tempfile

        tmp_target = Path(tempfile.mkdtemp(prefix="gigachat_ca_")) / MERGED_CA_BUNDLE_FILENAME
        tmp_target.write_text(content, encoding="utf-8")
        log.info("SSL: склейка сертификатов записана во временную папку: %s", tmp_target)
        return tmp_target
    except Exception:
        return None


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
        "interception": detect_ssl_interception(),
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
    """
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
            "GIGACHAT_TIMEOUT=30",
            "GIGACHAT_VERIFY_SSL_CERTS=true",
            "GIGACHAT_CA_BUNDLE_FILE=",
        ]
    elif not any(line.startswith("GIGACHAT_CA_BUNDLE_FILE") for line in lines):
        # Сохраняем настройку SSL, даже если её не было в старом файле
        lines.append("GIGACHAT_CA_BUNDLE_FILE=")

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
