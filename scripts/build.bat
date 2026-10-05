@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

:: Корень проекта = папка уровнем выше папки scripts
set "PROJECT_ROOT=%~dp0.."

echo ============================================================
echo 🚛 УСТАНОВКА ГЕНЕРАТОРА КП
echo ============================================================
echo.
echo Устанавливаю PyInstaller...
pip install pyinstaller
echo.
echo Запускаю сборку...
python "%~dp0setup.py"
echo.
pause
