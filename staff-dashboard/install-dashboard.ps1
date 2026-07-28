param(
    [switch]$NoAutoStart
)

$ErrorActionPreference = 'Stop'
$sourceExe = Join-Path $PSScriptRoot '청년이봄 직원알림.exe'
if (-not (Test-Path -LiteralPath $sourceExe)) {
    throw '청년이봄 직원알림.exe 파일을 설치 스크립트와 같은 폴더에 두어야 합니다.'
}

$installDirectory = Join-Path $env:LOCALAPPDATA 'YouthSpringStaffDashboard\App'
$installedExe = Join-Path $installDirectory '청년이봄 직원알림.exe'
New-Item -ItemType Directory -Force -Path $installDirectory | Out-Null

Get-Process -ErrorAction SilentlyContinue |
    Where-Object {
        try { $_.Path -eq $installedExe } catch { $false }
    } |
    ForEach-Object {
        if (-not $_.CloseMainWindow()) {
            Stop-Process -Id $_.Id -Force
        } else {
            [void]$_.WaitForExit(5000)
            if (-not $_.HasExited) {
                Stop-Process -Id $_.Id -Force
            }
        }
        try { [void]$_.WaitForExit(5000) } catch {}
    }

for ($attempt = 1; $attempt -le 5; $attempt++) {
    try {
        Copy-Item -LiteralPath $sourceExe -Destination $installedExe -Force
        break
    } catch [System.IO.IOException] {
        if ($attempt -eq 5) { throw }
        Start-Sleep -Milliseconds 500
    }
}

$shell = New-Object -ComObject WScript.Shell
$desktopShortcut = $shell.CreateShortcut(
    (Join-Path ([Environment]::GetFolderPath('Desktop')) '청년이봄 직원알림.lnk')
)
$desktopShortcut.TargetPath = $installedExe
$desktopShortcut.WorkingDirectory = $installDirectory
$desktopShortcut.Description = '청년이봄 현장예약 알림 및 운영 현황'
$desktopShortcut.Save()

if (-not $NoAutoStart) {
    $startupDirectory = [Environment]::GetFolderPath('Startup')
    $startupShortcut = $shell.CreateShortcut(
        (Join-Path $startupDirectory '청년이봄 직원알림.lnk')
    )
    $startupShortcut.TargetPath = $installedExe
    $startupShortcut.WorkingDirectory = $installDirectory
    $startupShortcut.Description = '청년이봄 현장예약 알림 자동 실행'
    $startupShortcut.Save()
}

Start-Process -FilePath $installedExe
Write-Host '설치를 완료했습니다.'
