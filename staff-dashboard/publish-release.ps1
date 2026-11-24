# 새 버전을 staff-dashboard\release에 만든다.
# 설치된 직원 프로그램은 10분 안에 이 파일을 찾아 스스로 업데이트한다(만들고 2분 뒤부터).
# 처음 설치하는 PC는 release\설치하기.cmd를 한 번 실행하면 된다.
$ErrorActionPreference = 'Stop'
$projectDirectory = $PSScriptRoot
$releaseDirectory = Join-Path $projectDirectory 'release'
$buildDirectory = Join-Path $projectDirectory 'publish-release-build'
$exeName = '청년이봄 직원알림.exe'

if (Test-Path -LiteralPath $buildDirectory) { Remove-Item -LiteralPath $buildDirectory -Recurse -Force }
dotnet publish (Join-Path $projectDirectory 'YouthSpring.StaffDashboard.csproj') `
    -c Release -r win-x64 --self-contained `
    -p:PublishSingleFile=true `
    --configfile (Join-Path $projectDirectory 'NuGet.Online.Config') `
    -o $buildDirectory
if ($LASTEXITCODE -ne 0) { throw '빌드에 실패했습니다.' }

$test = Start-Process -FilePath (Join-Path $buildDirectory $exeName) -ArgumentList '--self-test' -Wait -PassThru
if ($test.ExitCode -ne 0) { throw "자체 테스트에 실패했습니다(종료 코드 $($test.ExitCode)). release는 바꾸지 않았습니다." }

New-Item -ItemType Directory -Force -Path $releaseDirectory | Out-Null
foreach ($script in @('install-dashboard.ps1', 'uninstall-dashboard.ps1', '설치하기.cmd', '제거하기.cmd')) {
    Copy-Item -LiteralPath (Join-Path $projectDirectory $script) -Destination $releaseDirectory -Force
}
# 임시 이름으로 복사한 뒤 바꿔 넣어, 프로그램이 복사 중인 파일을 집어 가지 않게 한다.
$staging = Join-Path $releaseDirectory "$exeName.new"
Copy-Item -LiteralPath (Join-Path $buildDirectory $exeName) -Destination $staging -Force
(Get-Item -LiteralPath $staging).LastWriteTimeUtc = [DateTime]::UtcNow
Move-Item -LiteralPath $staging -Destination (Join-Path $releaseDirectory $exeName) -Force
Remove-Item -LiteralPath $buildDirectory -Recurse -Force

Write-Host "새 버전을 만들었습니다: $releaseDirectory"
Write-Host '설치된 직원 프로그램이 약 2~12분 안에 스스로 업데이트합니다.'
