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
echo   [1] Запустить генератор (шаблоны, без ИИ)
echo   [2] Запустить генератор (GigaChat - облачный ИИ)
echo   [3] Установить библиотеки
echo   [4] Ввести API-ключ GigaChat
echo   [5] Создать ярлык на рабочем столе
echo   [6] Очистить логи
echo   [7] Выход
echo.
echo ════════════════════════════════════════════
echo.
set /p choice="Выберите действие (1-7): "

if "%choice%"=="1" goto RUN_OLLAMA
if "%choice%"=="2" goto RUN_GIGACHAT
if "%choice%"=="3" goto INSTALL
if "%choice%"=="4" goto SET_KEY
if "%choice%"=="5" goto SHORTCUT
if "%choice%"=="6" goto CLEAR
if "%choice%"=="7" goto EXIT
goto MENU

:: ============================================
:: ЗАПУСК ГЕНЕРАТОРА (ШАБЛОНЫ, БЕЗ ИИ)
:: ============================================
:RUN_OLLAMA
cls
echo.
echo ╔═══════════════════════════════════════════╗
echo ║    ЗАПУСК ГЕНЕРАТОРА (ШАБЛОНЫ, БЕЗ ИИ)    ║
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
    echo Чтобы установить библиотеки, выберите в меню пункт 3
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
:: ЗАПУСК ГЕНЕРАТОРА (GIGACHAT)
:: ============================================
:RUN_GIGACHAT
cls
echo.
echo ╔═══════════════════════════════════════════╗
echo ║       ЗАПУСК ГЕНЕРАТОРА (GIGACHAT)       ║
echo ╚═══════════════════════════════════════════╝
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python не установлен!
    echo    Скачайте: https://www.python.org/downloads/
    echo.
    pause
    goto MENU
)

if not exist "%PROJECT_ROOT%\.env" (
    echo ⚠️  Файл .env не найден!
    echo    Сначала введите ключ (пункт 4 меню).
    echo.
    pause
    goto MENU
)

findstr /b "GIGACHAT_CREDENTIALS=" "%PROJECT_ROOT%\.env" | findstr /v "GIGACHAT_CREDENTIALS=$" >nul
if errorlevel 1 (
    echo ⚠️  Ключ GigaChat не задан!
    echo    Введите ключ (пункт 4 меню).
    echo.
    pause
    goto MENU
)

echo ✅ Ключ найден, библиотеки проверяются...
python -c "import gigachat, dotenv" >nul 2>&1
if errorlevel 1 (
    echo.
    echo ⚠️  Библиотеки GigaChat не установлены!
    echo    Установите их: pip install gigachat python-dotenv
    echo    Или выберите пункт 3 меню.
    echo.
    pause
    goto MENU
)

echo ⏳ Запуск...
echo.
echo ════════════════════════════════════════════
echo.

python "%PROJECT_ROOT%\src\генератор_кп_gigachat.py"

if errorlevel 1 (
    echo.
    echo ❌ Программа завершилась с ошибкой.
    echo    Проверьте логи в папке logs/
    echo.
    pause
)

goto MENU

:: ============================================
:: ВВОД API-КЛЮЧА GIGACHAT
:: ============================================
:SET_KEY
cls
echo.
echo ╔═══════════════════════════════════════════╗
echo ║         ВВОД API-КЛЮЧА GIGACHAT          ║
echo ╚═══════════════════════════════════════════╝
echo.
echo 🔑 Получите ключ: https://developers.sber.ru/
echo    (Личный кабинет -^> Настройки -^> Authorization Key)
echo.
echo ⚠️  Ключ сохранится в .env и НЕ попадёт в Git.
echo    Файл .env добавлен в .gitignore.
echo.

if not exist "%PROJECT_ROOT%\.env" (
    echo ⏳ Создаю .env с настройками по умолчанию...
    (
        echo GIGACHAT_CREDENTIALS=
        echo GIGACHAT_SCOPE=GIGACHAT_API_PERS
        echo GIGACHAT_MODEL=GigaChat
        echo GIGACHAT_TIMEOUT=30
        echo GIGACHAT_VERIFY_SSL_CERTS=true
        echo GIGACHAT_CA_BUNDLE_FILE=
    ) > "%PROJECT_ROOT%\.env"
)

set "key="
set /p "key=Вставьте Authorization Key (base64): "

if "!key!"=="" (
    echo.
    echo ❌ Ключ не введён. Отмена.
    echo.
    pause
    goto MENU
)

REM Удаляем старую строку с ключом и добавляем новую
findstr /v /b "GIGACHAT_CREDENTIALS" "%PROJECT_ROOT%\.env" > "%PROJECT_ROOT%\.env.tmp"
echo GIGACHAT_CREDENTIALS=!key!>> "%PROJECT_ROOT%\.env.tmp"
move /y "%PROJECT_ROOT%\.env.tmp" "%PROJECT_ROOT%\.env" >nul

echo.
echo ✅ Ключ сохранён в .env
echo 🔒 Проверьте, что .env в .gitignore (он там по умолчанию)
echo.
pause
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
pip install python-docx pyperclip requests gigachat python-dotenv
echo.

echo ⏳ Проверка установленных библиотек...
echo.
python -c "import docx; print('✅ python-docx установлен')" 2>nul || echo ❌ python-docx НЕ УСТАНОВЛЕН
python -c "import pyperclip; print('✅ pyperclip установлен')" 2>nul || echo ❌ pyperclip НЕ УСТАНОВЛЕН
python -c "import tkinter; print('✅ tkinter установлен')" 2>nul || echo ❌ tkinter НЕ УСТАНОВЛЕН
python -c "import requests; print('✅ requests установлен')" 2>nul || echo ❌ requests НЕ УСТАНОВЛЕН
python -c "import gigachat; print('✅ gigachat установлен')" 2>nul || echo ❌ gigachat НЕ УСТАНОВЛЕН
python -c "import dotenv; print('✅ python-dotenv установлен')" 2>nul || echo ❌ python-dotenv НЕ УСТАНОВЛЕН

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