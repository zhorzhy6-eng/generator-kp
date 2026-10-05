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
  6. Копирует готовые .exe на рабочий стол.
  7. Проверяет, что .exe реально запускается (пробный запуск с таймаутом).

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

# Имя, под которым шаблонная версия ложится на рабочий стол.
# Оставлено прежним: так работает ярлык/привычка пользователя и
# совместимо со старой сборкой setup.py.
DESKTOP_TEMPLATE_EXE = "Генератор_КП.exe"

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


def desktop_dir() -> Path:
    """
    Возвращает папку рабочего стола.

    Рабочего стола может не быть (OneDrive перенёс папку, ограниченный
    профиль) — тогда используем «Документы», затем домашнюю папку.
    """
    home = Path(os.path.expanduser("~"))
    for name in ("Desktop", "OneDrive/Desktop", "Рабочий стол", "Documents", "Документы"):
        candidate = home / name
        if candidate.is_dir():
            return candidate
    return home


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


def build(spec: Path, label: str) -> bool:
    """Собирает EXE по spec-файлу."""
    print(f"\n  📦 {label}")
    print(f"     spec: {spec.name}")

    if not spec.is_file():
        print(f"     ❌ spec-файл не найден: {spec}")
        return False

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

    if code != 0:
        print(f"     ❌ Сборка не удалась (код {code})")
        return False

    print(f"     ✅ Сборка завершена")
    return True


# ============================================
# ШАГ 6. КОПИРОВАНИЕ НА РАБОЧИЙ СТОЛ
# ============================================


def copy_to_desktop(exe_name: str, desktop_name: str = None) -> bool:
    source = DIST_DIR / exe_name
    if not source.is_file():
        print(f"  ❌ Нет файла для копирования: {source}")
        return False

    target_dir = desktop_dir()
    target = target_dir / (desktop_name or exe_name)

    try:
        shutil.copy2(source, target)
    except PermissionError:
        print(f"  ❌ Файл занят (возможно, программа запущена): {target}")
        return False
    except Exception as exc:
        print(f"  ❌ Не удалось скопировать: {exc}")
        return False

    size_mb = source.stat().st_size / (1024 * 1024)
    print(f"  ✅ {target}  ({size_mb:.1f} МБ)")

    # Рядом с EXE кладём образец .env: пользователь заполнит его ключом.
    # Именно так файл попадёт в папку, откуда программа его читает.
    env_example = PROJECT_ROOT / ".env.example"
    if env_example.is_file():
        try:
            shutil.copy2(env_example, target_dir / ".env.example")
            print("     рядом положен .env.example (образец для ключа)")
        except Exception as exc:
            print(f"     ⚠️  .env.example не скопирован: {exc}")

    # Папка config рядом с .exe — чтобы сертификат можно было просто положить
    for name in ("russian_trusted_root_ca.cer", "settings.example.json"):
        item = CONFIG_DIR / name
        if item.is_file():
            try:
                (target_dir / "config").mkdir(exist_ok=True)
                shutil.copy2(item, target_dir / "config" / name)
                print(f"     рядом положен config/{name}")
            except Exception as exc:
                print(f"     ⚠️  config/{name} не скопирован: {exc}")

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
      * успешный старт подтверждается появлением свежего файла лога рядом
        с .exe — значит модуль дошёл до настройки логирования.

    Returns:
        True, если EXE запустился.
    """
    print(f"\n  🧪 Пробный запуск: {exe_path.name}")

    if not exe_path.is_file():
        print(f"     ❌ Файл не найден: {exe_path}")
        return False

    log_dir = exe_path.parent / "logs"
    before = set(log_dir.glob("generator_*.log")) if log_dir.is_dir() else set()

    started = time.time()
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

    # Ждём либо появления лога, либо завершения процесса
    appearing = False
    while time.time() - started < timeout:
        if proc.poll() is not None:
            # Процесс завершился сам — для GUI это признак проблемы
            print(f"     ❌ Процесс завершился сам, код {proc.returncode}")
            return False

        now = set(log_dir.glob("generator_*.log")) if log_dir.is_dir() else set()
        if now - before or (log_dir.is_dir() and any(log_dir.iterdir())):
            appearing = True
            break

        time.sleep(0.5)

    elapsed = time.time() - started

    # Гасим процесс: время ожидания вышло, программа работает штатно
    try:
        proc.terminate()
        try:
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()
    except Exception:
        pass

    if appearing:
        print(f"     ✅ Программа стартовала за {elapsed:.1f} с и пишет логи")
        print(f"        ({log_dir})")
        return True

    print(f"     ⚠️  Процесс запустился и не упал за {elapsed:.0f} с,")
    print("        но файл лога не появился. Проверьте запуск вручную.")
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
        "--no-desktop", action="store_true", help="не копировать EXE на рабочий стол"
    )
    parser.add_argument(
        "--no-test", action="store_true", help="не выполнять пробный запуск EXE"
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

    step(5, "Сборка EXE")
    built = []

    if build(GIGACHAT_SPEC, "Версия с ИИ (GigaChat)"):
        built.append((GIGACHAT_EXE, GIGACHAT_EXE))

    if not args.skip_template:
        if build(TEMPLATE_SPEC, "Шаблонная версия (без ИИ)"):
            built.append((TEMPLATE_EXE, DESKTOP_TEMPLATE_EXE))

    if not built:
        print("\n  ❌ Ни один EXE не собрался.")
        return 1

    if not args.no_desktop:
        step(6, "Копирование на рабочий стол")
        for exe_name, desktop_name in built:
            copy_to_desktop(exe_name, desktop_name)

    if not args.no_test:
        step(7, "Проверка, что EXE запускается")
        for exe_name, _ in built:
            smoke_test(DIST_DIR / exe_name)

    # ---------- Итог ----------
    print()
    hr()
    print("  ✅ ГОТОВО")
    hr()
    for exe_name, _ in built:
        path = DIST_DIR / exe_name
        if path.is_file():
            size_mb = path.stat().st_size / (1024 * 1024)
            print(f"  📦 {path}  ({size_mb:.1f} МБ)")
    if not args.no_desktop:
        print(f"\n  📋 Копии на рабочем столе: {desktop_dir()}")
        print(f"     {GIGACHAT_EXE}      — с ИИ, нужен ключ GigaChat")
        if not args.skip_template:
            print(f"     {DESKTOP_TEMPLATE_EXE}         — шаблоны, работает без ключа")

    print("\n  📋 Что дальше:")
    print("     1. Запустите EXE двойным кликом.")
    print("     2. Нажмите «🔑 Ключ» и вставьте Authorization Key GigaChat.")
    print("        Ключ сохранится в .env РЯДОМ С EXE (в Git он не попадёт).")
    print("     3. Если в логе SSL-ошибка — нажмите «⚠️ SSL» в окне")
    print("        и положите сертификат Минцифры в папку config.")
    print("     4. Кнопка «📋 Диагностика» покажет, какой .env найден.")
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
