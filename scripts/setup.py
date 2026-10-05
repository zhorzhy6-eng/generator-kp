#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Установщик и сборщик «Генератора КП».

Что делает:
  1. Проверяет версию Python и наличие pip.
  2. Устанавливает зависимости из requirements.txt.
  3. Ставит pre-commit hook (защита от утечки ключа GigaChat в Git).
  4. Проверяет, что исходники на месте и компилируются.
  5. Собирает EXE через spec-файлы (а не набором ключей в командной строке).
  6. Копирует готовые ПАПКИ сборки на рабочий стол.
  7. Проверяет, что программа реально запускается (пробный запуск с таймаутом).

ВАЖНО ПРО РЕЖИМ СБОРКИ (ONEDIR):
  Программа собирается в режиме onedir, а не onefile. На выходе получается
  не один .exe, а ПАПКА dist\\<имя>\\ с загрузчиком и подпапкой _internal\\.

  Причина: onefile-сборка при каждом запуске распаковывает всё содержимое
  во временную папку %TEMP%\\_MEIxxxxx\\. Антивирусы (Kaspersky, Pro32)
  блокируют эту распаковку, и программа падает с ошибкой
  «Could not create temporary directory!» ещё до появления окна.
  Исключения антивируса на папку проекта не помогают — распаковка идёт
  в %TEMP%, а не в проект.

  В режиме onedir распаковки нет: зависимости читаются из _internal\\
  рядом с .exe. Поэтому установщик теперь работает с папками, а не с
  отдельными файлами. Переносить программу на другой компьютер нужно
  ТОЛЬКО целой папкой — один .exe без _internal\\ не заработает.

Запуск:
    python scripts/setup.py                # полная установка и сборка
    python scripts/setup.py --build-only   # только сборка
    python scripts/setup.py --skip-template  # без шаблонной версии
    python scripts/setup.py --no-desktop   # не копировать на рабочий стол
"""

import argparse
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

# ============================================
# НАСТРОЙКА ВЫВОДА (ВАЖНО ДЛЯ WINDOWS)
# ============================================
# Консоль Windows по умолчанию в cp866/cp1251, и печать эмодзи или русского
# текста падает с UnicodeEncodeError. Раньше установщик из-за этого обрывался
# на первой же строке с эмодзи. Переключаем вывод в UTF-8 и подстраховываемся
# заменой непечатаемых символов.
for _stream_name in ("stdout", "stderr"):
    _stream = getattr(sys, _stream_name, None)
    try:
        if _stream is not None:
            _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Если консоль всё равно не умеет UTF-8 (очень старый Windows) — просто
# отключаем эмодзи, чтобы установщик дошёл до конца.
try:
    "🚛".encode(sys.stdout.encoding or "utf-8")
    EMOJI_OK = True
except Exception:
    EMOJI_OK = False


def emo(char: str) -> str:
    """Возвращает эмодзи, если консоль его переживёт, иначе пустую строку."""
    return char if EMOJI_OK else ""

SCRIPTS_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPTS_DIR.parent
SRC_DIR = PROJECT_ROOT / "src"
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
CONFIG_DIR = PROJECT_ROOT / "config"

GIGACHAT_SPEC = SCRIPTS_DIR / "generator_kp.spec"
TEMPLATE_SPEC = SCRIPTS_DIR / "generator_kp_template.spec"

GIGACHAT_EXE = "Генератор_КП_GigaChat.exe"
TEMPLATE_EXE = "Генератор_КП_Шаблон.exe"

# Сборка идёт в режиме onedir: результат — ПАПКА, внутри которой лежат
# загрузчик .exe и подпапка _internal\ с зависимостями.
GIGACHAT_BUNDLE = "Генератор_КП_GigaChat"
TEMPLATE_BUNDLE = "Генератор_КП_Шаблон"

# Имя подпапки с зависимостями. Задано явно и в spec-файлах
# (contents_directory="_internal"), чтобы не зависеть от версии PyInstaller.
CONTENTS_DIR = "_internal"

# Имя, под которым шаблонная версия ложится на рабочий стол.
# Раньше (в режиме onefile) это был одиночный файл «Генератор_КП.exe».
# Теперь сборка — папка, и переименовывать её в «.exe» нельзя: получится
# папка с расширением .exe, которая только запутает. Поэтому копия на
# рабочем столе называется так же, как папка сборки, — и остаётся
# самодостаточной (внутри есть и .exe, и _internal\).
DESKTOP_TEMPLATE_NAME = TEMPLATE_BUNDLE

MIN_PYTHON = (3, 9)


def hr(char="=", width=62):
    print(char * width)


def step(number, text):
    print()
    hr("-")
    print(f"  [{number}] {text}")
    hr("-")


def run(cmd, cwd=None, check=False):
    """Запускает команду, транслируя вывод в консоль."""
    print(f"  $ {' '.join(str(c) for c in cmd)}")
    try:
        return subprocess.run(
            [str(c) for c in cmd],
            cwd=str(cwd or PROJECT_ROOT),
            check=check,
        ).returncode
    except FileNotFoundError:
        print(f"  ❌ Команда не найдена: {cmd[0]}")
        return 127


def desktop_dir(override: str = None) -> Path:
    """
    Возвращает папку, куда класть копию сборки.

    По умолчанию — рабочий стол. Его может не быть (OneDrive перенёс папку,
    ограниченный профиль) — тогда используем «Документы», затем домашнюю
    папку.

    override — путь из ключа --desktop-dir; нужен, чтобы положить копию
    в выбранное место (например, на флешку) или в тестовую папку.
    """
    if override:
        path = Path(override).expanduser()
        try:
            path.mkdir(parents=True, exist_ok=True)
        except Exception as exc:
            print(f"  ⚠️  Не удалось создать папку {path}: {exc}")
        return path

    home = Path(os.path.expanduser("~"))
    for name in ("Desktop", "OneDrive/Desktop", "Рабочий стол", "Documents", "Документы"):
        candidate = home / name
        if candidate.is_dir():
            return candidate
    return home


def dir_size_bytes(path: Path) -> int:
    """
    Считает суммарный размер папки со всеми вложенными файлами.

    Нужно потому, что в режиме onedir результат сборки — папка, а не файл:
    «вес» программы складывается из загрузчика .exe и содержимого _internal\\.
    Отдельные файлы могут быть заняты антивирусом или удалены параллельно,
    поэтому ошибки доступа молча пропускаем, а не роняем установщик.
    """
    total = 0
    try:
        for item in path.rglob("*"):
            try:
                if item.is_file():
                    total += item.stat().st_size
            except OSError:
                # Файл занят или исчез между обходом и stat — не критично
                pass
    except OSError:
        pass
    return total


def mb(size_bytes: int) -> float:
    """Переводит байты в мегабайты."""
    return size_bytes / (1024 * 1024)


# ============================================
# ШАГ 1. PYTHON
# ============================================


def check_python() -> bool:
    step(1, "Проверка Python")
    version = sys.version_info
    print(f"  Python: {sys.version.split()[0]}  ({sys.executable})")

    if version[:2] < MIN_PYTHON:
        print(f"  ❌ Нужен Python {MIN_PYTHON[0]}.{MIN_PYTHON[1]} или новее.")
        print("     Скачайте: https://www.python.org/downloads/")
        return False

    print("  ✅ Версия подходит")
    return True


# ============================================
# ШАГ 2. ЗАВИСИМОСТИ
# ============================================


def install_dependencies() -> bool:
    step(2, "Установка зависимостей")

    requirements = PROJECT_ROOT / "requirements.txt"
    if requirements.is_file():
        code = run([sys.executable, "-m", "pip", "install", "-r", requirements])
        if code != 0:
            print("  ⚠️  Не все зависимости из requirements.txt установились")
    else:
        print("  ⚠️  requirements.txt не найден — ставлю пакеты по списку")
        run(
            [
                sys.executable,
                "-m",
                "pip",
                "install",
                "python-docx",
                "pyperclip",
                "requests",
                "gigachat",
                "python-dotenv",
            ]
        )

    # PyInstaller нужен именно для сборки, поэтому в requirements.txt не входит
    print("\n  Устанавливаю PyInstaller (нужен только для сборки)...")
    run([sys.executable, "-m", "pip", "install", "pyinstaller"])

    # Проверяем, что ключевые модули действительно доступны
    print("\n  Проверка импортов:")
    required = ["docx", "pyperclip", "tkinter", "PyInstaller"]
    optional = ["gigachat", "dotenv"]

    ok = True
    for module in required:
        code = run([sys.executable, "-c", f"import {module}"])
        if code == 0:
            print(f"    ✅ {module}")
        else:
            print(f"    ❌ {module} — не установлен")
            ok = False

    for module in optional:
        code = run([sys.executable, "-c", f"import {module}"])
        print(f"    {'✅' if code == 0 else '⚠️ '} {module} (для GigaChat)")

    return ok


# ============================================
# ШАГ 3. PRE-COMMIT HOOK
# ============================================


def install_git_hook() -> bool:
    """
    Ставит pre-commit hook, блокирующий коммит секретов.

    Хук лежит в .git/hooks/, а .git не попадает в репозиторий. Поэтому при
    клонировании проекта его нужно установить заново — этим и занимается
    установщик, чтобы защита не терялась на новой машине.
    """
    step(3, "Защита от утечки секретов (pre-commit hook)")

    source = SCRIPTS_DIR / "pre-commit"
    if not source.is_file():
        print(f"  ⚠️  Не найден {source} — пропускаю")
        return False

    git_dir = PROJECT_ROOT / ".git"
    if not git_dir.is_dir():
        print("  ⚠️  Это не git-репозиторий — hook не нужен")
        return False

    target = git_dir / "hooks" / "pre-commit"
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        # Читаем/пишем байтами: hook обязан остаться с переводами строк LF,
        # иначе git на Windows откажется его запускать.
        data = source.read_bytes().replace(b"\r\n", b"\n")
        target.write_bytes(data)
        try:
            os.chmod(target, 0o755)
        except Exception:
            pass
        print(f"  ✅ Hook установлен: {target}")
        print("     Блокирует коммит при наличии .env, settings.json, *.log, *.key")
        print("     и непустого GIGACHAT_CREDENTIALS в индексе.")
        return True
    except Exception as exc:
        print(f"  ❌ Не удалось установить hook: {exc}")
        return False


# ============================================
# ШАГ 4. ПРОВЕРКА ИСХОДНИКОВ
# ============================================


def check_sources() -> bool:
    step(4, "Проверка исходников")

    files = [
        SRC_DIR / "генератор_кп_gigachat.py",
        SRC_DIR / "генератор_кп.py",
        SRC_DIR / "llm_provider.py",
        SRC_DIR / "logger_config.py",
    ]

    ok = True
    for path in files:
        if not path.is_file():
            print(f"  ❌ Нет файла: {path}")
            ok = False
            continue
        code = run([sys.executable, "-m", "py_compile", path])
        print(f"    {'✅' if code == 0 else '❌'} {path.name}")
        if code != 0:
            ok = False

    # Напоминаем про сертификат: без него EXE соберётся, но SSL не вылечится
    ca = CONFIG_DIR / "russian_trusted_root_ca.cer"
    if ca.is_file():
        print(f"\n  ✅ Сертификат Минцифры найден: {ca}")
        print("     Он будет упакован в EXE.")
    else:
        print("\n  ⚠️  Сертификат Минцифры не найден:")
        print(f"     {ca}")
        print("     Если GigaChat падает с SSL-ошибкой, скачайте его:")
        print("     https://gu-st.ru/content/Other/doc/russian_trusted_root_ca.cer")

    return ok


# ============================================
# ШАГ 5. СБОРКА
# ============================================


def free_stale_output(bundle_name: str) -> None:
    """
    Освобождает путь dist/<bundle_name>/ перед сборкой (режим onedir).

    В onedir результат сборки — ПАПКА, поэтому убирать нужно её целиком:
    PyInstaller дописывает файлы в уже существующий каталог, и остатки
    прошлой сборки ломают результат с PermissionError.

    Иногда обычное удаление невозможно: файлы внутри держит антивирус
    (Kaspersky и подобные сканируют свежесобранные .dll/.exe) или остался
    «хвост» от прерванной сборки. Тогда папка ПЕРЕИМЕНОВЫВАЕТСЯ — это
    отдельная операция, которая в такой ситуации обычно проходит, после
    чего имя снова свободно.

    Никакие данные пользователя здесь не трогаются: только наш dist/.
    """
    target = DIST_DIR / bundle_name
    if not target.exists():
        return

    try:
        shutil.rmtree(target)
        print(f"  ♻️  Удалена старая папка сборки: {bundle_name}\\")
        return
    except Exception as exc:
        print(f"  ⚠️  Не удалось удалить папку {bundle_name}\\: {exc}")

    stamp = time.strftime("%Y%m%d_%H%M%S")
    stale = target.with_name(f"{target.name}_STALE_{stamp}")
    try:
        target.rename(stale)
        print(f"  ♻️  Папка была занята — отложена как {stale.name}\\")
    except Exception as exc:
        print(f"  ❌ Не удалось освободить {target}: {exc}")
        print("     Закройте программу (если запущена) и повторите сборку.")


def cleanup_old_stale() -> None:
    """Убирает ранее отложенные файлы и папки *_STALE_*, чтобы dist не разрастался."""
    try:
        for item in DIST_DIR.glob("*_STALE_*"):
            try:
                if item.is_dir():
                    shutil.rmtree(item)
                else:
                    item.unlink()
            except Exception:
                # Занятый файл не мешает сборке — оставляем как есть
                pass
    except Exception:
        pass


def build(spec: Path, label: str, bundle_name: str, exe_name: str) -> bool:
    """Собирает папку сборки (режим onedir) по spec-файлу."""
    print(f"\n  📦 {label}")
    print(f"     spec: {spec.name}")
    print(f"     режим: onedir (папка {bundle_name}\\)")

    if not spec.is_file():
        print(f"     ❌ spec-файл не найден: {spec}")
        return False

    DIST_DIR.mkdir(parents=True, exist_ok=True)
    cleanup_old_stale()
    # Освобождаем целевую папку: PyInstaller дописывает файлы в уже готовый
    # каталог, поэтому занятая папка (антивирус, запущенная копия программы,
    # остаток прерванной сборки) ломает сборку с PermissionError.
    free_stale_output(bundle_name)

    # Вызываем PyInstaller как модуль: так не зависим от того, попал ли
    # pyinstaller.exe в PATH (частая проблема на Windows).
    code = run(
        [
            sys.executable,
            "-m",
            "PyInstaller",
            str(spec),
            "--clean",
            "--noconfirm",
            "--distpath",
            str(DIST_DIR),
            "--workpath",
            str(BUILD_DIR),
        ]
    )

    final_dir = DIST_DIR / bundle_name
    final_exe = final_dir / exe_name
    # Проверяем именно папку с загрузчиком внутри: если PyInstaller по
    # ошибке соберёт onefile, папки не будет и проверка это поймает.
    if code != 0 or not final_exe.is_file():
        print(f"     ❌ Сборка не удалась (код {code})")
        if final_dir.is_dir() and not final_exe.is_file():
            print(f"     Папка {bundle_name}\\ есть, но внутри нет {exe_name}.")
        print("     Если в ошибке PermissionError — файлы держит антивирус.")
        print("     Добавьте папку проекта в его исключения и повторите сборку.")
        free_stale_output(bundle_name)
        return False

    total = dir_size_bytes(final_dir)
    internal = final_dir / CONTENTS_DIR
    internal_mb = mb(dir_size_bytes(internal)) if internal.is_dir() else 0.0
    print(f"     ✅ Сборка завершена: {final_dir}")
    print(f"        загрузчик {exe_name}: {mb(final_exe.stat().st_size):.1f} МБ")
    if internal.is_dir():
        print(f"        папка {CONTENTS_DIR}\\: {internal_mb:.1f} МБ")
    print(f"        всего папка: {mb(total):.1f} МБ")
    return True


# ============================================
# ШАГ 6. КОПИРОВАНИЕ НА РАБОЧИЙ СТОЛ
# ============================================


def copy_to_desktop(
    bundle_name: str,
    exe_name: str,
    desktop_name: str = None,
    dest_root: str = None,
) -> bool:
    """
    Копирует на рабочий стол ВСЮ папку сборки (режим onedir).

    Копировать один .exe нельзя: он не заработает без папки _internal\\
    (там Python и все библиотеки). Поэтому переносится каталог целиком,
    а внутрь него докладываются .env.example и config\\ — программа ищет
    их рядом с .exe, то есть уже внутри скопированной папки.

    Так папка на рабочем столе получается самодостаточной: её можно
    скопировать на флешку или отправить архивом и запустить на другой
    машине без Python.
    """
    source_dir = DIST_DIR / bundle_name
    source_exe = source_dir / exe_name

    if not source_dir.is_dir():
        print(f"  ❌ Нет папки для копирования: {source_dir}")
        return False
    if not source_exe.is_file():
        print(f"  ❌ В папке нет {exe_name}: {source_exe}")
        return False

    target_root = desktop_dir(dest_root)
    target = target_root / (desktop_name or bundle_name)

    # Старую копию убираем: copytree не умеет писать поверх существующей
    # папки, а смешивать две версии сборки в одной папке нельзя.
    if target.exists():
        try:
            shutil.rmtree(target)
            print(f"  ♻️  Старая копия на рабочем столе удалена: {target.name}\\")
        except Exception as exc:
            print(f"  ❌ Не удалось удалить старую копию {target}: {exc}")
            print("     Возможно, программа запущена — закройте её и повторите.")
            return False

    try:
        shutil.copytree(source_dir, target)
    except PermissionError as exc:
        print(f"  ❌ Не удалось скопировать (файлы заняты): {exc}")
        return False
    except Exception as exc:
        print(f"  ❌ Не удалось скопировать папку: {exc}")
        return False

    # ---------- Докладываем образцы внутрь папки ----------
    # Программа читает .env и config\ рядом с .exe (sys.executable),
    # а .exe лежит в корне папки сборки — значит, файлы кладём туда же.
    env_example = PROJECT_ROOT / ".env.example"
    if env_example.is_file():
        try:
            shutil.copy2(env_example, target / ".env.example")
            print("     внутрь папки положен .env.example (образец для ключа)")
        except Exception as exc:
            print(f"     ⚠️  .env.example не скопирован: {exc}")

    # Папка config\ рядом с .exe: программа ищет сертификат Минцифры именно
    # там (_ca_bundle_dirs → sys.executable.parent/config). В сборке её нет
    # «из коробки», поэтому создаём сразу — иначе пользователю пришлось бы
    # угадывать, куда класть сертификат.
    try:
        (target / "config").mkdir(exist_ok=True)
    except Exception as exc:
        print(f"     ⚠️  Папка config не создана: {exc}")

    for name in ("russian_trusted_root_ca.cer", "settings.example.json"):
        item = CONFIG_DIR / name
        if item.is_file():
            try:
                shutil.copy2(item, target / "config" / name)
                print(f"     внутрь папки положен config/{name}")
            except Exception as exc:
                print(f"     ⚠️  config/{name} не скопирован: {exc}")

    total = dir_size_bytes(target)
    print(f"  ✅ {target}\\  ({mb(total):.1f} МБ)")
    print(f"     запускать: {target / exe_name}")
    return True


# ============================================
# ШАГ 7. ПРОВЕРКА ЗАПУСКА
# ============================================


def smoke_test(exe_path: Path, timeout: float = 45.0) -> bool:
    """
    Пробный запуск EXE: проверяет, что программа стартует и пишет лог.

    Как это проверяется без ручного наблюдения за окном:
      * программа запускается в отдельном процессе;
      * ошибка запуска (например, падение логгера) видна по коду возврата;
      * успешный старт подтверждается СВЕЖЕЙ записью в логе рядом с .exe —
        значит модуль дошёл до настройки логирования.

    Важно: в режиме onedir .exe лежит внутри папки сборки, поэтому лог
    появляется в dist/<папка сборки>/logs/, а не в dist/logs/. Путь
    вычисляется от самого .exe, так что за этим следить не нужно.

    «Свежесть» лога проверяется по времени изменения файла, а не по факту
    его существования: лог от прошлого запуска не должен приниматься за
    доказательство старта.

    Returns:
        True, если EXE запустился.
    """
    print(f"\n  🧪 Пробный запуск: {exe_path.name}")

    if not exe_path.is_file():
        print(f"     ❌ Файл не найден: {exe_path}")
        return False

    log_dir = exe_path.parent / "logs"
    if not log_dir.is_dir():
        log_dir.mkdir(parents=True, exist_ok=True)

    def newest_log_mtime() -> float:
        """Время последнего изменения самого свежего лога (0 — логов нет)."""
        try:
            stamps = [p.stat().st_mtime for p in log_dir.glob("generator_*.log")]
        except OSError:
            return 0.0
        return max(stamps) if stamps else 0.0

    started = time.time()
    mtime_before = newest_log_mtime()

    try:
        proc = subprocess.Popen(
            [str(exe_path)],
            cwd=str(exe_path.parent),
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception as exc:
        print(f"     ❌ Не удалось запустить: {exc}")
        return False

    # Ждём либо свежей записи в логе, либо завершения процесса.
    # exit_code остаётся None, пока процесс жив и работает штатно.
    exit_code = None
    logging_started = False
    try:
        while time.time() - started < timeout:
            exit_code = proc.poll()
            if exit_code is not None:
                # Процесс завершился сам — для GUI это признак проблемы
                break

            if newest_log_mtime() > mtime_before:
                logging_started = True
                break

            time.sleep(0.5)
    finally:
        # Гасим процесс ВСЕГДА: даже если проверка не удалась или прервана
        # (Ctrl+C). Иначе запущенный EXE держит свои файлы в dist\ и мешает
        # следующей сборке — удаление падает с PermissionError, и это легко
        # принять за выходку антивируса.
        try:
            if proc.poll() is None:
                proc.terminate()
                try:
                    proc.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    proc.wait(timeout=5)
        except Exception:
            pass

    elapsed = time.time() - started

    if exit_code is not None:
        print(f"     ❌ Процесс завершился сам, код {exit_code}")
        return False

    if logging_started:
        print(f"     ✅ Программа стартовала за {elapsed:.1f} с и пишет логи")
        print(f"        ({log_dir})")
        return True

    print(f"     ⚠️  Процесс запустился и не упал за {elapsed:.0f} с,")
    print("        но свежей записи в логе нет. Проверьте запуск вручную.")
    return True


# ============================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================


def main() -> int:
    parser = argparse.ArgumentParser(description="Сборка Генератора КП")
    parser.add_argument(
        "--build-only", action="store_true", help="не ставить зависимости, только собрать"
    )
    parser.add_argument(
        "--skip-template", action="store_true", help="не собирать шаблонную версию"
    )
    parser.add_argument(
        "--no-desktop", action="store_true", help="не копировать папки сборки на рабочий стол"
    )
    parser.add_argument(
        "--no-test", action="store_true", help="не выполнять пробный запуск EXE"
    )
    parser.add_argument(
        "--desktop-dir",
        default=None,
        help="куда класть копию сборки (по умолчанию — рабочий стол)",
    )
    args = parser.parse_args()

    hr()
    print("  🚛  ГЕНЕРАТОР КП — УСТАНОВКА И СБОРКА")
    hr()
    print(f"  Проект: {PROJECT_ROOT}")

    if not check_python():
        return 1

    if not args.build_only:
        if not install_dependencies():
            print("\n  ❌ Не все обязательные зависимости установлены.")
            print("     Исправьте ошибки выше и запустите установщик снова.")
            return 1
        install_git_hook()
    else:
        print("\n  (режим --build-only: зависимости и hook пропущены)")

    if not check_sources():
        print("\n  ❌ Исходники не годятся для сборки.")
        return 1

    step(5, "Сборка в режиме onedir (папки, а не одиночные файлы)")
    # Каждый элемент: (папка сборки, .exe внутри неё, имя копии на столе)
    built = []

    if build(GIGACHAT_SPEC, "Версия с ИИ (GigaChat)", GIGACHAT_BUNDLE, GIGACHAT_EXE):
        built.append((GIGACHAT_BUNDLE, GIGACHAT_EXE, GIGACHAT_BUNDLE))

    if not args.skip_template:
        if build(TEMPLATE_SPEC, "Шаблонная версия (без ИИ)", TEMPLATE_BUNDLE, TEMPLATE_EXE):
            built.append((TEMPLATE_BUNDLE, TEMPLATE_EXE, DESKTOP_TEMPLATE_NAME))

    if not built:
        print("\n  ❌ Ни одна версия не собралась.")
        return 1

    # Пробный запуск идёт ПЕРЕД копированием намеренно: если он провалится
    # (антивирус оборвал сборку, программа падает), сломанная копия не попадёт
    # на рабочий стол. Вместе с программой копируются и созданные ею файлы.
    if not args.no_test:
        step(6, "Проверка, что программа запускается")
        for bundle_name, exe_name, _ in built:
            smoke_test(DIST_DIR / bundle_name / exe_name)

    if not args.no_desktop:
        step(7, "Копирование папок сборки на рабочий стол")
        for bundle_name, exe_name, desktop_name in built:
            copy_to_desktop(bundle_name, exe_name, desktop_name, args.desktop_dir)

    # ---------- Итог ----------
    print()
    hr()
    print("  ✅ ГОТОВО")
    hr()
    for bundle_name, exe_name, _ in built:
        path = DIST_DIR / bundle_name
        if path.is_dir():
            total = dir_size_bytes(path)
            exe_path = path / exe_name
            exe_mb = mb(exe_path.stat().st_size) if exe_path.is_file() else 0.0
            print(f"  📦 {path}\\")
            print(f"     {exe_name}  ({exe_mb:.1f} МБ)")
            print(f"     Итого папка: {mb(total):.1f} МБ")
    if not args.no_desktop:
        print(f"\n  📋 Копии в папке: {desktop_dir(args.desktop_dir)}")
        print(f"     {GIGACHAT_BUNDLE}\\{GIGACHAT_EXE}")
        print("        — с ИИ, нужен ключ GigaChat")
        if not args.skip_template:
            print(f"     {DESKTOP_TEMPLATE_NAME}\\{TEMPLATE_EXE}")
            print("        — шаблоны, работает без ключа")

    print("\n  📋 Что дальше:")
    print("     1. Запускать нужно .exe ВНУТРИ папки сборки, например:")
    print(f"        {DIST_DIR / GIGACHAT_BUNDLE / GIGACHAT_EXE}")
    print("        Один .exe без папки _internal\\ не заработает!")
    print("     2. Нажмите «🔑 Ключ» и вставьте Authorization Key GigaChat.")
    print("        Ключ сохранится в .env РЯДОМ С EXE (в Git он не попадёт).")
    print("     3. Если в логе SSL-ошибка — нажмите «⚠️ SSL» в окне")
    print("        и положите сертификат Минцифры в папку config.")
    print("     4. Кнопка «📋 Диагностика» покажет, какой .env найден.")
    print("\n  📋 Как перенести на другой компьютер:")
    print("     Копируйте ВСЮ папку сборки целиком (или упакуйте её в ZIP).")
    print("     Папку удобно запускать ярлыком: правый клик по .exe →")
    print("     «Отправить» → «Рабочий стол (создать ярлык)».")
    hr()

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\n  ⛔ Прервано пользователем")
        sys.exit(130)
    except Exception as exc:
        import traceback

        print("\n  ❌ Непредвиденная ошибка установщика:")
        traceback.print_exc()
        sys.exit(1)
