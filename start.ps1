# RTF 转换服务启动脚本
# 功能：PID 防重复启动 + 端口占用检测清理 + 后台静默启动 + 就绪探测
param(
    [string]$Python = "",
    [int]$Port = 8000
)

$ErrorActionPreference = "Stop"
$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PidFile = Join-Path $BaseDir "server.pid"
$LogFile = Join-Path $BaseDir "server.log"
$Url = "http://127.0.0.1:$Port"

# 1. 防重复启动（PID 文件）
if (Test-Path $PidFile) {
    $oldPid = Get-Content $PidFile -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($oldPid) {
        $proc = Get-Process -Id $oldPid -ErrorAction SilentlyContinue
        if ($proc -and $proc.ProcessName -match "python") {
            Write-Host "[提示] 服务已在运行 (PID=$oldPid)，如需重启请先执行 stop"
            exit 1
        }
    }
}

# 2. 端口占用检测与清理（仅清理本服务的 python 进程）
$conn = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
if ($conn) {
    $pid_ = $conn.OwningProcess
    $p = Get-Process -Id $pid_ -ErrorAction SilentlyContinue
    if ($p -and $p.ProcessName -match "python") {
        Stop-Process -Id $pid_ -Force
        Write-Host "[清理] 端口 $Port 被 python 占用，已结束 PID=$pid_"
    } else {
        Write-Host "[错误] 端口 $Port 被其他进程占用 PID=$pid_ 名称=$($p.ProcessName)"
        exit 1
    }
} else {
    Write-Host "[检测] 端口 $Port 空闲"
}

# 3. 自动查找 Python（优先级：命令行指定 > 已知路径 > py launcher > PATH，自动排除沙箱路径）
function Resolve-Python {
    param([string]$Hint)
    $candidates = @()
    if ($Hint -and (Test-Path $Hint)) { $candidates += $Hint }

    # ========== 修改这里：优先当前项目 .venv 虚拟环境 ==========
    $venvPython = Join-Path $BaseDir ".venv\Scripts\python.exe"
    $known = @(
        $venvPython,
        "D:\python\python.exe"
    )
    # ========================================================

    foreach ($p in $known) { if (Test-Path $p) { $candidates += $p } }
    # py launcher
    try {
        $pyPath = & py -3 -c "import sys; print(sys.executable)" 2>$null
        if ($pyPath -and (Test-Path $pyPath)) { $candidates += $pyPath }
    } catch {}
    # PATH 里的 python 最后（可能是沙箱/其他不相关 Python）
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd) { $candidates += $cmd.Source }
    # 去重 + 排除沙箱路径 + 验证能跑
    $seen = @{}
    foreach ($c in $candidates) {
        if (-not $c -or $seen.ContainsKey($c)) { continue }
        $seen[$c] = $true
        if ($c -match "Doubao" -or $c -match "sandbox_runtime") { continue }
        try {
            & $c --version 2>$null | Out-Null
            if ($LASTEXITCODE -eq 0) { return $c }
        } catch {}
    }
    return $null
}


$Python = Resolve-Python -Hint $Python
if (-not $Python) {
    Write-Host "[错误] 找不到 Python。请安装 Python 3.10+（勾选 Add to PATH），或用 -Python 参数指定 python.exe 完整路径"
    exit 1
}
Write-Host "[Python] $Python"

# 3.5 首次运行自动安装依赖（flask / pywin32；tkinter 是标准库自带）
$needInstall = $false
try {
    & $Python -c "import flask, win32com" 2>$null | Out-Null
    if ($LASTEXITCODE -ne 0) { $needInstall = $true }
} catch { $needInstall = $true }
if ($needInstall) {
    Write-Host "[依赖] 首次运行，自动安装 flask / pywin32 ..."
    try { & $Python -m pip install flask pywin32 --quiet } catch { Write-Host "[警告] 自动装依赖失败，请手动执行: pip install flask pywin32" }
}

# 4. 后台静默启动（无窗口残留）
$p = Start-Process -FilePath $Python -ArgumentList "app.py" -WorkingDirectory $BaseDir -WindowStyle Hidden -PassThru
Set-Content -Path $PidFile -Value $p.Id -Encoding ascii
Write-Host "[已启动] PID=$($p.Id)"

# 5. 等待服务就绪
$ok = $false
for ($i = 0; $i -lt 15; $i++) {
    try {
        Invoke-RestMethod -Uri "$Url/api/health" -TimeoutSec 2 | Out-Null
        $ok = $true
        break
    } catch { Start-Sleep -Seconds 1 }
}
if ($ok) { Write-Host "[就绪] 服务已就绪" } else { Write-Host "[警告] 15 秒内未检测到服务，请查看 $LogFile" }

Write-Host ""
Write-Host "服务地址:   $Url"
Write-Host "健康检查:   $Url/api/health"
Write-Host "停止服务:   stop.ps1"
