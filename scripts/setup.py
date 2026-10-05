import os
import shutil
import subprocess
from pathlib import Path

print("=" * 60)
print("🚛 УСТАНОВЩИК ГЕНЕРАТОРА КП")
print("=" * 60)

# Пути
# Этот файл лежит в <корень проекта>/scripts/, исходники — в <корень проекта>/src/
scripts_dir = Path(__file__).resolve().parent
project_root = scripts_dir.parent
source_file = project_root / 'src' / 'генератор_кп.py'
desktop = Path(os.path.expanduser('~')) / 'Desktop'

# Сборка (запускается из корня проекта, чтобы build/ и dist/ создавались там же)
print("\n📦 Собираю программу...")
subprocess.run(
    [
        'pyinstaller', '--onefile', '--console',
        '--name', 'Генератор_КП',
        '--distpath', str(project_root / 'dist'),
        '--workpath', str(project_root / 'build'),
        '--specpath', str(project_root),
        str(source_file),
    ],
    cwd=str(project_root),
    check=False,
)

# Копируем на рабочий стол
exe_file = project_root / 'dist' / 'Генератор_КП.exe'
if exe_file.exists():
    shutil.copy(exe_file, desktop / 'Генератор_КП.exe')
    print("\n✅ ГОТОВО! Программа на рабочем столе!")
else:
    print("\n❌ Ошибка! Файл не создан.")
    print(f"   Ожидался файл: {exe_file}")
    print("   Проверьте, что PyInstaller установлен: pip install pyinstaller")

print("=" * 60)
input("\nНажмите Enter для выхода...")
