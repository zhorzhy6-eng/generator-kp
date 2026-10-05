# -*- mode: python ; coding: utf-8 -*-
# ============================================================
#  PyInstaller spec: «Генератор КП» с ИИ (GigaChat)
# ============================================================
#  Сборка выполняется ТОЛЬКО через этот spec-файл, а не набором ключей
#  в командной строке: так надёжнее и воспроизводимо.
#
#  Запуск (из корня проекта):
#      pyinstaller scripts/generator_kp.spec --clean --noconfirm
#
#  Результат: dist/Генератор_КП_GigaChat.exe
#
#  Рассчитан на запуск двойным кликом на машине БЕЗ установленного Python:
#  все зависимости упаковываются внутрь одного файла.
# ============================================================

import os
import sys

from PyInstaller.utils.hooks import collect_submodules

# ---------- Корень проекта ----------
# spec-файл лежит в <корень>/scripts/, поэтому корень — на уровень выше.
# SPECPATH PyInstaller выставляет в папку spec-файла.
try:
    PROJECT_ROOT = os.path.dirname(os.path.abspath(SPECPATH))
except NameError:  # pragma: no cover
    PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(os.getcwd())))

SRC_DIR = os.path.join(PROJECT_ROOT, "src")
CONFIG_DIR = os.path.join(PROJECT_ROOT, "config")

MAIN_SCRIPT = os.path.join(SRC_DIR, "генератор_кп_gigachat.py")

if not os.path.isfile(MAIN_SCRIPT):
    raise SystemExit(f"❌ Не найден главный файл: {MAIN_SCRIPT}")

# ---------- Скрытые импорты ----------
# collect_submodules нужен потому, что gigachat/dotenv/docx подгружают часть
# модулей динамически — PyInstaller их «не видит» при статическом анализе.
hiddenimports = []
for package in ("gigachat", "dotenv", "docx", "httpx", "pydantic"):
    try:
        hiddenimports += collect_submodules(package)
    except Exception:
        pass

hiddenimports += [
    "pyperclip",
    "tkinter",
    "tkinter.ttk",
    "tkinter.messagebox",
    "tkinter.filedialog",
    "tkinter.scrolledtext",
    "logger_config",
    "llm_provider",
]

# ---------- Данные ----------
# ВАЖНО: .env (с настоящим ключом) НИКОГДА не упаковывается в .exe.
# Кладём только .env.example как образец.
datas = []

env_example = os.path.join(PROJECT_ROOT, ".env.example")
if os.path.isfile(env_example):
    datas.append((env_example, "."))

settings_example = os.path.join(CONFIG_DIR, "settings.example.json")
if os.path.isfile(settings_example):
    datas.append((settings_example, "config"))

# Корневой сертификат Минцифры — если пользователь его уже скачал
ca_bundle = os.path.join(CONFIG_DIR, "russian_trusted_root_ca.cer")
if os.path.isfile(ca_bundle):
    datas.append((ca_bundle, "config"))
    print(f"[spec] Сертификат Минцифры будет упакован: {ca_bundle}")
else:
    print(
        "[spec] ВНИМАНИЕ: config/russian_trusted_root_ca.cer не найден.\n"
        "       EXE соберётся, но при SSL-ошибке сертификат нужно будет\n"
        "       положить рядом с .exe в папку config."
    )

icon_path = os.path.join(CONFIG_DIR, "icon.ico")
has_icon = os.path.isfile(icon_path)

print(f"[spec] Корень проекта: {PROJECT_ROOT}")
print(f"[spec] Главный файл:   {MAIN_SCRIPT}")
print(f"[spec] Скрытых импортов: {len(hiddenimports)}")

# ---------- Анализ ----------
a = Analysis(
    [MAIN_SCRIPT],
    pathex=[SRC_DIR],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Тяжёлые библиотеки не нужны: уменьшают размер и ускоряют сборку
    excludes=[
        "matplotlib",
        "numpy",
        "pandas",
        "scipy",
        "IPython",
        "pytest",
        "notebook",
        "PyQt5",
        "PySide2",
        "PySide6",
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Генератор_КП_GigaChat",
    # suffix='.bin' — намеренно НЕ .exe.
    #
    # Причина: PyInstaller пишет итоговый файл в несколько приёмов (копирует
    # bootloader, встраивает манифест, дописывает архив). Антивирусы и
    # защитные механизмы Windows начинают сканировать свежий .exe сразу после
    # появления и держат его на запись — тогда последний шаг сборки падает с
    # PermissionError. Файл с расширением .bin под это не попадает, а имя
    # меняется на .exe последним шагом уже в setup.py.
    suffix=".bin",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    # console=False — без чёрного окна консоли.
    # Логи всё равно пишутся в файл (logs/ рядом с .exe), а также доступны
    # кнопкой «📋 Диагностика» в интерфейсе.
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path if has_icon else None,
)
