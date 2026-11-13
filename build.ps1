# 나이스 성적 분석 - 빌드 스크립트
# 사용법: PowerShell에서  .\build.ps1
# 결과물: dist\나이스성적분석.exe  (이 파일 하나만 배포하면 됩니다)
#         + OneDrive 작업 폴더의 완성본\ 에도 복사합니다.

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$dotnet = (Get-Command dotnet -ErrorAction SilentlyContinue).Source
if (-not $dotnet) { $dotnet = Join-Path $env:LOCALAPPDATA "Microsoft\dotnet\dotnet.exe" }
if (-not (Test-Path $dotnet)) { throw ".NET SDK를 찾을 수 없습니다. https://dot.net 에서 SDK를 설치하세요." }

$env:DOTNET_CLI_TELEMETRY_OPTOUT = "1"
$env:DOTNET_NOLOGO = "1"

$proj = Join-Path $root "src\NeisAnalyzer\NeisAnalyzer.csproj"
$out = Join-Path $root "build-output"
$dist = Join-Path $root "dist"

Write-Host "빌드 중..." -ForegroundColor Cyan
& $dotnet build $proj -c Release -o $out --nologo -v quiet
if ($LASTEXITCODE -ne 0) { throw "빌드 실패" }

New-Item -ItemType Directory -Force $dist | Out-Null
Copy-Item (Join-Path $out "나이스성적분석.exe") (Join-Path $dist "나이스성적분석.exe") -Force
$size = (Get-Item (Join-Path $dist "나이스성적분석.exe")).Length
Write-Host ("완료: dist\나이스성적분석.exe  ({0:N0} KB)" -f ($size / 1KB)) -ForegroundColor Green

# OneDrive 작업 폴더의 완성본 폴더에도 복사 (있을 때만)
$final = Join-Path ([Environment]::GetFolderPath("UserProfile")) "OneDrive - 거원중학교\성적분석\완성본"
if (Test-Path (Split-Path $final -Parent)) {
    New-Item -ItemType Directory -Force $final | Out-Null
    try {
        Copy-Item (Join-Path $dist "나이스성적분석.exe") (Join-Path $final "나이스성적분석.exe") -Force
        Write-Host "복사: $final\나이스성적분석.exe" -ForegroundColor Green
    } catch {
        Write-Host "완성본 폴더로 복사하지 못했습니다(OneDrive가 파일을 잠갔을 수 있음). 잠시 후 다시 .\build.ps1 을 실행하세요." -ForegroundColor Yellow
    }
}
