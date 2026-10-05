@echo off
rem ============================================================
rem  Обёртка для Планировщика задач Windows.
rem  Запускает автосинхронизацию проекта с GitHub.
rem  Обычная консоль cmd: работает без Store-алиасов и без входа в систему.
rem ============================================================
chcp 65001 >nul
set "PROJECT_ROOT=%~dp0.."
pwsh -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0sync-to-github.ps1" -RepoPath "%PROJECT_ROOT%"
exit /b %errorlevel%
