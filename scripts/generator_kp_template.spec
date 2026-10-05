# -*- mode: python ; coding: utf-8 -*-
# ============================================================
#  PyInstaller spec: «Генератор КП» (шаблонная версия, без ИИ)
# ============================================================
#  Работает БЕЗ ключа GigaChat и без интернета: текст коммерческого
#  предложения собирается из шаблонных фраз. Нужна как запасной вариант
#  и для машин, где ИИ недоступен.
#
#  Запуск (из корня проекта):
#      pyinstaller scripts/generator_kp_template.spec --clean --noconfirm
#
#  Результат (режим ONEDIR): папка
#      dist/Генератор_КП_Шаблон/
#      ├── Генератор_КП_Шаблон.exe   — загрузчик
#      ├── _internal/                — Python и все библиотеки
#      └── config/
#
#  Почему НЕ onefile — см. подробный комментарий в generator_kp.spec:
#  onefile распаковывается в %TEMP%\_MEIxxxxx\ при каждом запуске, и
#  антивирус (Kaspersky, Pro32) блокирует эту распаковку. В onedir
#  распаковки нет, зависимости читаются из _internal\ рядом с .exe.
#
#  GigaChat и python-dotenv здесь НЕ нужны — поэтому они исключены:
#  папка получается заметно меньше.
# ============================================================

import os

from PyInstaller.utils.hooks import collect_submodules

try:
    PROJECT_ROOT = os.path.dirname(os.path.abspath(SPECPATH))
except NameError:  # pragma: no cover
    PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(os.getcwd())))

SRC_DIR = os.path.join(PROJECT_ROOT, "src")
CONFIG_DIR = os.path.join(PROJECT_ROOT, "config")

MAIN_SCRIPT = os.path.join(SRC_DIR, "генератор_кп.py")

if not os.path.isfile(MAIN_SCRIPT):
    raise SystemExit(f"❌ Не найден главный файл: {MAIN_SCRIPT}")

hiddenimports = []
for package in ("docx",):
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
]

datas = []

settings_example = os.path.join(CONFIG_DIR, "settings.example.json")
if os.path.isfile(settings_example):
    datas.append((settings_example, "config"))

icon_path = os.path.join(CONFIG_DIR, "icon.ico")
has_icon = os.path.isfile(icon_path)

print(f"[spec] Корень проекта: {PROJECT_ROOT}")
print(f"[spec] Главный файл:   {MAIN_SCRIPT}")
print(f"[spec] Скрытых импортов: {len(hiddenimports)}")

a = Analysis(
    [MAIN_SCRIPT],
    pathex=[SRC_DIR],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # В шаблонной версии ИИ не используется — выкидываем всё его окружение
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
        "gigachat",
        "dotenv",
        "httpx",
        "requests",
        "pydantic",
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    # ONEDIR: библиотеки и данные уходят в COLLECT() ниже, в _internal\.
    # exclude_binaries=True — именно этот флаг переключает сборку в onedir.
    exclude_binaries=True,
    name="Генератор_КП_Шаблон",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path if has_icon else None,
)

# ---------- Сборка папки (ONEDIR) ----------
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="Генератор_КП_Шаблон",
    contents_directory="_internal",
)
