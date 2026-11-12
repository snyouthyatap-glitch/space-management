param(
    # 연결 설정(관리자 비밀번호·네이버 계정)과 알림 이력·오류 로그까지 지우려면 붙인다.
    [switch]$RemoveSettings
)

$ErrorActionPreference = 'Stop'
$dataDirectory = Join-Path $env:LOCALAPPDATA 'YouthSpringStaffDashboard'
$installDirectory = Join-Path $dataDirectory 'App'
$installedExe = Join-Path $installDirectory '청년이봄 직원알림.exe'

# 1. 실행 중인 직원 프로그램 끄기(트레이에 숨어 있는 것 포함)
Get-Process -ErrorAction SilentlyContinue |
    Where-Object {
        try { $_.Path -eq $installedExe } catch { $false }
    } |
    ForEach-Object {
        Stop-Process -Id $_.Id -Force
        try { [void]$_.WaitForExit(5000) } catch {}
    }

# 2. 바탕화면·시작프로그램 바로가기 지우기
$shortcuts = @(
    (Join-Path ([Environment]::GetFolderPath('Desktop')) '청년이봄 직원알림.lnk'),
    (Join-Path ([Environment]::GetFolderPath('Startup')) '청년이봄 직원알림.lnk')
)
foreach ($shortcut in $shortcuts) {
    if (Test-Path -LiteralPath $shortcut) {
        Remove-Item -LiteralPath $shortcut -Force
        Write-Host "바로가기 삭제: $shortcut"
    }
}

# 3. 프로그램 파일 지우기
if (Test-Path -LiteralPath $installDirectory) {
    Remove-Item -LiteralPath $installDirectory -Recurse -Force
    Write-Host "프로그램 삭제: $installDirectory"
}

# 4. 설정은 기본으로 남긴다(다시 설치하면 비밀번호를 다시 넣지 않아도 된다).
if ($RemoveSettings -and (Test-Path -LiteralPath $dataDirectory)) {
    Remove-Item -LiteralPath $dataDirectory -Recurse -Force
    Write-Host "설정·기록 삭제: $dataDirectory"
} elseif (Test-Path -LiteralPath $dataDirectory) {
    Write-Host "설정·알림 이력·오류 로그는 남겨 두었습니다: $dataDirectory"
}

Write-Host '제거를 완료했습니다.'
