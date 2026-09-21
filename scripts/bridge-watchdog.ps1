# Сторож моста «Hermes для Excel»: если мост не отвечает — поднимает его.
#
# Зачем: мост — обычный локальный процесс, он не должен переживать перезагрузку «на честном слове».
# Сторож проверяет /health и запускает мост заново, если тот упал. Он вызывается
# планировщиком задач раз в минуту и работает только пока запущен Excel.
#
# Когда Excel закрыт, сторож останавливает мост. Это важно: процесс моста использует venv Hermes
# и иначе блокирует обновление Hermes на Windows. Всё пишется в workspace\watchdog.log.

param(
    [switch]$Force,          # поднимать мост, даже если Excel сейчас не запущен
    [int]$Port = 3443
)

$ErrorActionPreference = 'Continue'
$root   = Split-Path -Parent $PSScriptRoot
$log    = Join-Path $root 'workspace\watchdog.log'
$bridge = Join-Path $root 'bridge\bridge.py'

function Write-Log([string]$msg) {
    $line = "[{0}] {1}" -f (Get-Date -Format 'dd.MM HH:mm:ss'), $msg
    Add-Content -Path $log -Value $line -Encoding UTF8
}

function Test-Bridge {
    try {
        $r = Invoke-WebRequest -UseBasicParsing -TimeoutSec 5 "https://localhost:$Port/health"
        return ($r.StatusCode -eq 200)
    } catch {
        return $false
    }
}

function Resolve-Python {
    # ВАЖНО: в PATH на этой машине первым лежит заглушка Microsoft Store
    # (…\WindowsApps\python.exe) — она печатает «Python» и открывает Store вместо запуска.
    # Поэтому сначала пробуем интерпретатор, которым пользуется сам Hermes, и каждого
    # кандидата проверяем делом: он должен реально выполнить код и вернуть версию.
    $cands = @()
    $cands += (Join-Path $env:LOCALAPPDATA 'hermes\hermes-agent\venv\Scripts\python.exe')
    $cands += (Join-Path $env:USERPROFILE 'hermes-excel\.venv\Scripts\python.exe')
    $fromPath = (Get-Command python -ErrorAction SilentlyContinue).Source
    if ($fromPath -and $fromPath -notlike '*WindowsApps*') { $cands += $fromPath }

    foreach ($c in $cands) {
        if (-not (Test-Path $c)) { continue }
        try {
            $v = & $c -c "import sys; print(sys.version.split()[0])" 2>$null
            if ($LASTEXITCODE -eq 0 -and $v) { return @{ exe = $c; pre = @() } }
        } catch { }
    }
    $pyLauncher = (Get-Command py -ErrorAction SilentlyContinue).Source
    if ($pyLauncher) {
        try {
            $v = & $pyLauncher -3 -c "import sys; print(sys.version.split()[0])" 2>$null
            if ($LASTEXITCODE -eq 0 -and $v) { return @{ exe = $pyLauncher; pre = @('-3') } }
        } catch { }
    }
    return $null
}

# Замок вместо именованного мьютекса: зависший запуск (или убитый) не блокирует сторожа навсегда —
# замок считается протухшим через 3 минуты. Найденная вживую беда: одна застрявшая копия держала
# мьютекс, и все следующие запуски молча выходили, оставляя мост лежать.
$lock = Join-Path $root 'workspace\watchdog.lock'
if (Test-Path $lock) {
    $age = (Get-Date) - (Get-Item $lock).LastWriteTime
    if ($age.TotalMinutes -lt 3) { exit 0 }   # прямо сейчас работает другой сторож
    Write-Log "протухший замок сторожа ($([int]$age.TotalMinutes) мин) — снимаю"
    Remove-Item $lock -Force -ErrorAction SilentlyContinue
}
Set-Content -Path $lock -Value (Get-Date -Format 's') -Encoding UTF8

try {
    $excel = Get-Process EXCEL -ErrorAction SilentlyContinue
    if (-not $excel -and -not $Force) {
        $holder = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
        if ($holder) {
            $pids = $holder | Select-Object -ExpandProperty OwningProcess | Sort-Object -Unique
            Write-Log ("Excel не запущен — останавливаю мост (PID $($pids -join ', '))")
            foreach ($p in $pids) { Stop-Process -Id $p -Force -ErrorAction SilentlyContinue }
        } else {
            Write-Log 'Excel не запущен — мост не поднимаю'
        }
        exit 0
    }

    if (Test-Bridge) { exit 0 }                # мост жив — ничего не делаем

    # Порт занят, но /health молчит — это зависший старый мост, его надо снять
    $holder = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
    if ($holder) {
        $pids = $holder | Select-Object -ExpandProperty OwningProcess | Sort-Object -Unique
        Write-Log ("порт $Port занят ($($pids -join ', ')), но /health не отвечает — останавливаю их")
        foreach ($p in $pids) { Stop-Process -Id $p -Force -ErrorAction SilentlyContinue }
        Start-Sleep -Seconds 2
    }

    $py = Resolve-Python
    if (-not $py) {
        Write-Log 'не нашёл рабочий python (заглушка WindowsApps не считается) — мост не запущен'
        exit 1
    }

    Write-Log "поднимаю мост: $($py.exe) $($py.pre -join ' ') $bridge"
    # Без перенаправления потоков: мост сам пишет workspace\bridge.log (см. bridge.py, log()).
    # RedirectStandardOutput здесь подвешивал сторожа, если дочерний процесс не отдавал поток.
    Start-Process -FilePath $py.exe -ArgumentList (@($py.pre) + @($bridge)) -WorkingDirectory $root `
        -WindowStyle Hidden

    for ($i = 0; $i -lt 20; $i++) {
        Start-Sleep -Milliseconds 700
        if (Test-Bridge) { Write-Log 'мост поднялся и отвечает на /health'; exit 0 }
    }
    Write-Log 'мост не ответил за 14 с — смотрите workspace\bridge.log'
    exit 1
} finally {
    Remove-Item $lock -Force -ErrorAction SilentlyContinue
}
