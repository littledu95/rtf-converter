# RTF 转换服务停止脚本
$ErrorActionPreference = "SilentlyContinue"
$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PidFile = Join-Path $BaseDir "server.pid"
$Port = 8000

if (-not (Test-Path $PidFile)) {
    Write-Host "[提示] 未找到 PID 文件，服务可能未运行"
    exit 0
}
$pid_ = Get-Content $PidFile | Select-Object -First 1
Stop-Process -Id $pid_ -Force
if ($?) {
    Write-Host "[已停止] PID=$pid_"
} else {
    Write-Host "[提示] 进程 $pid_ 已不存在"
}
Remove-Item $PidFile -Force
Write-Host "若端口仍被占用，可手动检查: netstat -ano | findstr :$Port"
