@echo off
rem ============================================================
rem  Повторный вход в GitHub (если автосинхронизация потеряла доступ).
rem  Открывает обычное окно Git Credential Manager: войдите в браузере,
rem  и автосинхронизация снова заработает.
rem ============================================================
chcp 65001 >nul
echo ============================================================
echo  ВХОД В GITHUB ДЛЯ АВТОСИНХРОНИЗАЦИИ
echo ============================================================
echo.
echo Сейчас откроется окно входа. Выберите "Sign in with your browser".
echo.

rem Сбрасываем прежние сохранённые данные и переменные неинтерактивного режима
set "GCM_INTERACTIVE="
set "GCM_GITHUB_AUTHMODE="
set "GIT_TERMINAL_PROMPT=1"

pushd "%~dp0.."
echo Выход из текущего аккаунта...
"C:\Program Files\Git\cmd\git.exe" credential-manager github logout 2>nul

echo Запрашиваю новый вход...
"C:\Program Files\Git\cmd\git.exe" push origin main
set "RC=%errorlevel%"
popd

echo.
if "%RC%"=="0" (
    echo ✅ Вход выполнен, автосинхронизация работает.
) else (
    echo ❌ Вход не выполнен. Код ошибки: %RC%
)
echo.
pause
