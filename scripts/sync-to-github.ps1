# ============================================================
#  Автосинхронизация проекта с GitHub
# ============================================================
#  Проверяет состояние репозитория, коммитит изменения и делает push.
#  Рассчитан на запуск по расписанию (Планировщик задач Windows),
#  поэтому внутри есть блокировка от одновременного запуска двух копий.
#
#  Ручной запуск:
#     pwsh -NoProfile -ExecutionPolicy Bypass -File "scripts\sync-to-github.ps1"
# ============================================================

param(
    [string]$RepoPath = (Split-Path -Parent $PSScriptRoot),
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# ---------- Блокировка параллельного запуска ----------
$mutex = New-Object System.Threading.Mutex($false, 'Global\DshKpSyncToGitHub')
if (-not $mutex.WaitOne(0)) {
    exit 0   # другая копия уже работает — выходим молча
}

try {
    $git = 'C:\Program Files\Git\cmd\git.exe'
    if (-not (Test-Path $git)) { $git = 'git' }

    $logDir = Join-Path $RepoPath 'logs'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $logFile = Join-Path $logDir ("sync_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd'))

    function Write-Log([string]$Message) {
        $line = "{0}  {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
        Add-Content -Path $logFile -Value $line -Encoding utf8
    }

    function Invoke-Git([string[]]$Arguments) {
        # Возвращает текст вывода; код возврата в $script:GitExitCode
        $output = & $git -C $RepoPath @Arguments 2>&1 | Out-String
        $script:GitExitCode = $LASTEXITCODE
        return $output.Trim()
    }

    # ---------- Проверяем, что это git-репозиторий ----------
    Invoke-Git @('rev-parse', '--is-inside-work-tree') | Out-Null
    if ($script:GitExitCode -ne 0) { Write-Log "ОШИБКА: $RepoPath не является git-репозиторием"; exit 1 }

    # ---------- Проверяем доступ на запись к GitHub ----------
    # Если авторизация потеряна, незачем плодить локальные коммиты: они будут
    # копиться молча. Лучше сообщить об этом в лог сразу.
    $env:GIT_TERMINAL_PROMPT = '0'
    Invoke-Git @('ls-remote', '--exit-code', 'origin', 'HEAD') | Out-Null
    if ($script:GitExitCode -ne 0) {
        Write-Log "ОШИБКА: нет доступа к origin (авторизация GitHub)."
        Write-Log "  Что делать: выполните в папке проекта  git push  и войдите в браузере заново,"
        Write-Log "  либо запустите: `"$PSScriptRoot\auth-github.bat`""
        exit 1
    }

    # ---------- Есть ли что коммитить ----------
    $status = Invoke-Git @('status', '--porcelain')
    if (-not $status -and -not $Force) {
        # Изменений нет. Проверим, не нужно ли отправить уже сделанные коммиты
        $ahead = Invoke-Git @('rev-list', '--count', '@{u}..HEAD')
        if ($script:GitExitCode -eq 0 -and [int]$ahead -gt 0) {
            Write-Log "Локальные коммиты впереди origin на $ahead — отправляю"
        } else {
            exit 0   # всё синхронизировано, тихо выходим
        }
    } else {
        # ---------- Коммитим ----------
        $changed = ($status -split "`n" | Where-Object { $_.Trim() }).Count
        Invoke-Git @('add', '-A') | Out-Null

        $staged = Invoke-Git @('diff', '--cached', '--name-only')
        if ($staged) {
            Write-Log "Изменений: $changed. Файлы: $($staged -replace "`n", ', ')"
            $message = "Автосохранение: {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            Invoke-Git @('commit', '-m', $message) | Out-Null
            if ($script:GitExitCode -ne 0) {
                Write-Log "ОШИБКА коммита (код $script:GitExitCode)"
                exit 1
            }
            Write-Log "Коммит создан: $message"
        } else {
            Write-Log "Индекс пуст после add (всё под .gitignore?) — пропускаю коммит"
        }
    }

    # ---------- Отправляем на GitHub ----------
    $env:GIT_TERMINAL_PROMPT = '0'
    $pushOut = Invoke-Git @('push', 'origin', 'main')
    if ($script:GitExitCode -eq 0) {
        $sha = Invoke-Git @('rev-parse', '--short', 'HEAD')
        Write-Log "PUSH OK -> origin/main ($sha)"
    } else {
        Write-Log "ОШИБКА PUSH (код $script:GitExitCode): $pushOut"
        exit 1
    }
}
finally {
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
