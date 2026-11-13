# 자동 시험: 가짜 예시 파일로 분석을 돌려 결과가 나오고 검산이 맞는지 확인한다.
# 사용법: PowerShell에서  .\tests\run-smoke.ps1
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$exe = Join-Path $root "dist\나이스성적분석.exe"
if (-not (Test-Path $exe)) { $exe = Join-Path $root "build-output\나이스성적분석.exe" }
if (-not (Test-Path $exe)) { throw "먼저 .\build.ps1 로 빌드하세요." }

$tmp = Join-Path $env:TEMP "neis-smoke"
New-Item -ItemType Directory -Force $tmp | Out-Null
$fail = 0
foreach ($case in @("sample", "sample-midterm")) {
    $in = Join-Path $root "tests\$case"
    $out = Join-Path $tmp "$case.xlsx"
    if (Test-Path $out) { Remove-Item $out -Force }
    & $exe --cli $in --out $out | Out-Host
    $code = $LASTEXITCODE
    if ($code -eq 0 -and (Test-Path $out)) { Write-Host "[통과] $case → $out" -ForegroundColor Green }
    else { Write-Host "[실패] $case (종료 코드 $code)" -ForegroundColor Red; $fail++ }
}
if ($fail -gt 0) { throw "실패 $fail건" }
Write-Host "모든 자동 시험 통과" -ForegroundColor Green
