@echo off
chcp 65001 >nul
title Генератор КП - Автоперевозки
setlocal enabledelayedexpansion

:: ============================================
:: УНИВЕРСАЛЬНЫЙ ЗАПУСК ГЕНЕРАТОРА КП
:: ============================================

:: Корень проекта = папка уровнем выше папки scripts
set "PROJECT_ROOT=%~dp0.."
:: Код проекта лежит в src\ ; добавляем его в PYTHONPATH, чтобы работал импорт logger_config
set "PYTHONPATH=%PROJECT_ROOT%\src;%PYTHONPATH%"

:: Работаем из корня проекта: сюда пишется settings.json и папка logs
cd /d "%PROJECT_ROOT%"

:MENU
cls
echo.
echo ╔═══════════════════════════════════════════╗
echo ║     ГЕНЕРАТОР КОММЕРЧЕСКИХ ПРЕДЛОЖЕНИЙ    ║
echo ╚═══════════════════════════════════════════╝
echo.
echo   [1] Запустить генератор
echo   [2] Установить библиотеки
echo   [3] Создать ярлык на рабочем столе
echo   [4] Очистить логи
echo   [5] Выход
echo.
echo ════════════════════════════════════════════
echo.
set /p choice="Выберите действие (1-5): "

if "%choice%"=="1" goto RUN
if "%choice%"=="2" goto INSTALL
if "%choice%"=="3" goto SHORTCUT
if "%choice%"=="4" goto CLEAR
if "%choice%"=="5" goto EXIT
goto MENU

:: ============================================
:: ЗАПУСК ГЕНЕРАТОРА
:: ============================================
:RUN
cls
echo.
echo ╔═══════════════════════════════════════════╗
echo ║            ЗАПУСК ГЕНЕРАТОРА             ║
echo ╚═══════════════════════════════════════════╝
echo.

REM Проверяем наличие Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ ОШИБКА: Python не установлен!
    echo.
    echo Скачайте Python с официального сайта:
    echo https://www.python.org/downloads/
    echo.
    pause
    goto MENU
)

REM Проверяем наличие библиотек
echo ⏳ Проверка библиотек...
python -c "import docx, pyperclip, tkinter" >nul 2>&1
if errorlevel 1 (
    echo.
    echo ⚠️ ВНИМАНИЕ: Не все библиотеки установлены!
    echo.
    echo Чтобы установить библиотеки, выберите в меню пункт 2
    echo.
    pause
    goto MENU
)

echo ✅ Все библиотеки установлены!
echo.
echo ⏳ Запуск генератора...
echo.
echo ════════════════════════════════════════════
echo.

python "%PROJECT_ROOT%\src\генератор_кп.py"

if errorlevel 1 (
    echo.
    echo ❌ ОШИБКА: Программа завершилась с ошибкой!
    echo Проверьте файлы в папке logs для деталей.
    echo.
    pause
)

goto MENU

:: ============================================
:: УСТАНОВКА БИБЛИОТЕК
:: ============================================
:INSTALL
cls
echo.
echo ╔═══════════════════════════════════════════╗
echo ║        УСТАНОВКА БИБЛИОТЕК               ║
echo ╚═══════════════════════════════════════════╝
echo.

REM Проверяем наличие Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ ОШИБКА: Python не установлен!
    echo.
    echo Скачайте Python с официального сайта:
    echo https://www.python.org/downloads/
    echo.
    pause
    goto MENU
)

echo ✅ Python найден!
echo.
echo ⏳ Обновление pip...
python -m pip install --upgrade pip
echo.

echo ⏳ Установка библиотек...
pip install python-docx pyperclip
echo.

echo ⏳ Проверка установленных библиотек...
echo.
python -c "import docx; print('✅ python-docx установлен')" 2>nul || echo ❌ python-docx НЕ УСТАНОВЛЕН
python -c "import pyperclip; print('✅ pyperclip установлен')" 2>nul || echo ❌ pyperclip НЕ УСТАНОВЛЕН
python -c "import tkinter; print('✅ tkinter установлен')" 2>nul || echo ❌ tkinter НЕ УСТАНОВЛЕН

echo.
echo ✅ Установка завершена!
echo.
pause
goto MENU

:: ============================================
:: СОЗДАНИЕ ЯРЛЫКА
:: ============================================
:SHORTCUT
cls
echo.
echo ╔═══════════════════════════════════════════╗
echo ║        СОЗДАНИЕ ЯРЛЫКА НА РАБОЧЕМ СТОЛЕ  ║
echo ╚═══════════════════════════════════════════╝
echo.

set "SCRIPT_PATH=%~dp0Генератор КП.bat"
set "DESKTOP=%USERPROFILE%\Desktop"
set "SHORTCUT_NAME=Генератор КП"

echo ⏳ Создаю ярлык "Генератор КП" на рабочем столе...

REM Проверяем, существует ли уже ярлык
if exist "%DESKTOP%\%SHORTCUT_NAME%.lnk" (
    echo.
    echo ⚠️ Ярлык уже существует!
    set /p overwrite="Перезаписать? (Y/N): "
    if /i not "!overwrite!"=="Y" (
        echo ❌ Создание отменено.
        pause
        goto MENU
    )
    del "%DESKTOP%\%SHORTCUT_NAME%.lnk" 2>nul
)

REM Создаем VBS скрипт для создания ярлыка
(
echo Set oWS = WScript.CreateObject("WScript.Shell"^)
echo sLinkFile = "%DESKTOP%\%SHORTCUT_NAME%.lnk"
echo Set oLink = oWS.CreateShortcut(sLinkFile^)
echo oLink.TargetPath = "%SCRIPT_PATH%"
echo oLink.WorkingDirectory = "%PROJECT_ROOT%"
echo oLink.Description = "Генератор коммерческих предложений"
echo oLink.IconLocation = "shell32.dll, 14"
echo oLink.Save
) > "%TEMP%\create_shortcut.vbs"

cscript //nologo "%TEMP%\create_shortcut.vbs"
del "%TEMP%\create_shortcut.vbs"

if exist "%DESKTOP%\%SHORTCUT_NAME%.lnk" (
    echo.
    echo ✅ Ярлык успешно создан!
    echo 📁 Путь: %DESKTOP%\%SHORTCUT_NAME%.lnk
) else (
    echo.
    echo ❌ Не удалось создать ярлык.
)

echo.
pause
goto MENU

:: ============================================
:: ОЧИСТКА ЛОГОВ
:: ============================================
:CLEAR
cls
echo.
echo ╔═══════════════════════════════════════════╗
echo ║           ОЧИСТКА ФАЙЛОВ ЛОГОВ           ║
echo ╚═══════════════════════════════════════════╝
echo.

REM Старый единый лог в корне проекта (пишет генератор_кп.py)
if exist "%PROJECT_ROOT%\generator_kp.log" (
    echo ⏳ Удаляю файл логов generator_kp.log...
    del "%PROJECT_ROOT%\generator_kp.log"
    echo ✅ Лог-файл успешно удалён!
) else (
    echo ℹ️ Файл generator_kp.log не найден.
)

REM Логи с датой в папке logs (пишет logger_config.py)
if exist "%PROJECT_ROOT%\logs\*.log" (
    echo.
    echo ⏳ Удаляю логи из папки logs...
    del /q "%PROJECT_ROOT%\logs\*.log"
    echo ✅ Логи из папки logs удалены!
) else (
    echo ℹ️ Логи в папке logs не найдены.
)

if exist "%PROJECT_ROOT%\settings.json" (
    echo.
    set /p clear_settings="Очистить настройки (settings.json)? (Y/N): "
    if /i "!clear_settings!"=="Y" (
        del "%PROJECT_ROOT%\settings.json"
        echo ✅ Настройки удалены!
    )
)

echo.
pause
goto MENU

:: ============================================
:: ВЫХОД
:: ============================================
:EXIT
cls
echo.
echo ╔═══════════════════════════════════════════╗
echo ║            ДО СВИДАНИЯ!                  ║
echo ╚═══════════════════════════════════════════╝
echo.
timeout /t 2 /nobreak >nul
exit

:: ============================================
:: НАСТРОЙКА ДЛЯ РАБОТЫ С ПЕРЕМЕННЫМИ
:: ============================================
setlocal enabledelayedexpansion