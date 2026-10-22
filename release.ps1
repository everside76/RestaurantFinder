<#
.SYNOPSIS
  RestaurantFinder 릴리스 스크립트 (Windows PowerShell 5.1 호환)

.DESCRIPTION
  버전 교체 → 컴파일 확인 → PyInstaller 빌드 → 인증키 유출 검사 → 커밋 → 태그 → 푸시
  → 인증키 유출 재검사 → GitHub 릴리스 생성(RestaurantFinder.exe 첨부)
  각 단계가 실패하면 즉시 중단한다. 커밋 전에 실패하면 APP_VERSION 변경을 되돌린다.

.EXAMPLE
  .\release.ps1 -Version 1.2.3
  .\release.ps1 -Version 1.2.3 -Notes "검색 속도 개선, 오류 메시지 정리"
#>
param(
    [Parameter(Mandatory = $true)]
    [string]$Version,

    [string]$Notes = ""
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$MainFile = "restaurant_api_gui.py"
$ExePath  = "dist\RestaurantFinder.exe"
$Repo     = "everside76/RestaurantFinder"
$script:VersionChanged = $false
$script:Committed = $false
$script:NeedCommit = $true

function Write-Step([string]$Message) {
    Write-Host ""
    Write-Host "==> $Message" -ForegroundColor Cyan
}

function Stop-Release([string]$Message) {
    Write-Host ""
    Write-Host "[중단] $Message" -ForegroundColor Red
    # 커밋 전에 실패했다면 APP_VERSION 변경을 되돌린다
    if ($script:VersionChanged -and -not $script:Committed) {
        git checkout -- $MainFile
        if ($?) { Write-Host "  $MainFile 의 버전 변경을 되돌렸습니다." -ForegroundColor Yellow }
    }
    exit 1
}

# 네이티브 명령 실행 결과 확인 ($? 와 종료 코드 모두 검사)
function Assert-Native([bool]$Ok, [string]$What) {
    if (-not $Ok -or $LASTEXITCODE -ne 0) {
        Stop-Release "$What 실패 (종료 코드 $LASTEXITCODE)"
    }
}

# UTF-8(BOM 없음)으로 임시 파일 작성 — git/gh에 한글을 깨짐 없이 전달
function New-Utf8TempFile([string]$Content) {
    $path = [System.IO.Path]::GetTempFileName()
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($path, $Content, $enc)
    return $path
}

# ── 0. 입력/도구 확인 ─────────────────────────────
$Version = $Version.Trim().TrimStart('v', 'V')
if ($Version -notmatch '^\d+\.\d+\.\d+$') {
    Stop-Release "버전 형식이 올바르지 않습니다: '$Version' (예: 1.2.3)"
}
$Tag = "v$Version"

Write-Step "도구 확인 (git, python, gh)"
foreach ($tool in @("git", "python", "gh")) {
    if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) {
        Stop-Release "$tool 명령을 찾을 수 없습니다."
    }
}
gh auth status | Out-Null
Assert-Native $? "gh 로그인 확인(gh auth status)"

# ── 1. 작업 트리 확인 ─────────────────────────────
Write-Step "작업 트리 확인"
$dirty = git status --porcelain
Assert-Native $? "git status"
if ($dirty) {
    Write-Host ($dirty -join "`n")
    Stop-Release "커밋되지 않은 변경이 있습니다. 먼저 커밋하거나 정리하세요."
}
git rev-parse -q --verify "refs/tags/$Tag" | Out-Null
if ($LASTEXITCODE -eq 0) {
    Stop-Release "태그 $Tag 가 이미 존재합니다."
}
$global:LASTEXITCODE = 0

# ── 2. APP_VERSION 교체 ───────────────────────────
Write-Step "APP_VERSION → $Version"
$fullPath = Join-Path $PSScriptRoot $MainFile
$utf8 = New-Object System.Text.UTF8Encoding($false)
$src = [System.IO.File]::ReadAllText($fullPath, $utf8)
$pattern = '(?m)^APP_VERSION\s*=\s*"[^"]*"'
$found = [regex]::Matches($src, $pattern)
if ($found.Count -ne 1) {
    Stop-Release "APP_VERSION 정의를 정확히 1개 찾지 못했습니다 (찾은 개수: $($found.Count))."
}
$oldLine = $found[0].Value
$newSrc = [regex]::Replace($src, $pattern, "APP_VERSION = `"$Version`"")
if ($newSrc -eq $src) {
    # 첫 릴리스처럼 소스의 APP_VERSION 이 이미 목표 버전이면 교체·커밋 없이 현재 커밋(HEAD)을 그대로 릴리스한다
    # (작업 트리는 위에서 깨끗함을 확인했으므로 빌드물 = HEAD 소스)
    $script:NeedCommit = $false
    Write-Host "  APP_VERSION 이 이미 $Version 입니다 - 버전 교체·커밋 없이 현재 커밋으로 릴리스합니다."
} else {
    [System.IO.File]::WriteAllText($fullPath, $newSrc, $utf8)
    $script:VersionChanged = $true
    Write-Host "  $oldLine  →  APP_VERSION = `"$Version`""
}

# ── 3. 컴파일 확인 ────────────────────────────────
Write-Step "py_compile"
python -m py_compile $MainFile
Assert-Native $? "py_compile"

# ── 4. PyInstaller 빌드 ───────────────────────────
Write-Step "PyInstaller 빌드"
python -m PyInstaller --onefile --noconsole --clean --name "RestaurantFinder" --icon "app.ico" --add-data "app.ico;." $MainFile
Assert-Native $? "PyInstaller 빌드"
if (-not (Test-Path -LiteralPath $ExePath)) {
    Stop-Release "$ExePath 가 생성되지 않았습니다."
}

# ── 5. 빌드 부산물 정리 ───────────────────────────
Write-Step "build 폴더 / spec 파일 정리"
if (Test-Path -LiteralPath "build") { Remove-Item -LiteralPath "build" -Recurse -Force }
if (Test-Path -LiteralPath "RestaurantFinder.spec") { Remove-Item -LiteralPath "RestaurantFinder.spec" -Force }

# ── 6. 인증키 유출 검사 (커밋 전) ─────────────────
Write-Step "인증키 유출 검사 (커밋 전)"
python -I check_no_secrets.py
Assert-Native $? "인증키 유출 검사 — 커밋·태그·푸시·릴리스를 진행하지 않습니다"

# ── 7. 커밋 ───────────────────────────────────────
Write-Step "커밋"
if ($script:NeedCommit) {
    $msg = "릴리스 $Tag`n`nCo-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`n"
    $msgFile = New-Utf8TempFile $msg
    try {
        git commit -a -F $msgFile
        Assert-Native $? "git commit"
    } finally {
        Remove-Item -LiteralPath $msgFile -Force -ErrorAction SilentlyContinue
    }
} else {
    Write-Host "  변경 사항이 없어 커밋을 건너뜁니다 (현재 HEAD 를 릴리스)."
}
$script:Committed = $true

# ── 8. 태그 ───────────────────────────────────────
Write-Step "태그 $Tag"
git tag $Tag
Assert-Native $? "git tag"

# ── 9. 푸시 ───────────────────────────────────────
Write-Step "푸시 (브랜치 + 태그)"
git push
Assert-Native $? "git push"
git push origin $Tag
Assert-Native $? "git push origin $Tag"

# ── 10. 인증키 유출 재검사 (릴리스 직전) ──────────
Write-Step "인증키 유출 재검사 (릴리스 생성 직전)"
python -I check_no_secrets.py
Assert-Native $? "인증키 유출 재검사 — GitHub 릴리스를 생성하지 않습니다"

# ── 11. GitHub 릴리스 생성 ────────────────────────
Write-Step "GitHub 릴리스 생성"
if ([string]::IsNullOrWhiteSpace($Notes)) {
    $Notes = "RestaurantFinder $Tag 릴리스입니다.`n`n" +
             "- 이미 설치된 프로그램은 실행 시 자동으로 업데이트됩니다.`n" +
             "- 처음 사용한다면 아래 RestaurantFinder.exe 를 내려받아 실행하세요."
}
$notesFile = New-Utf8TempFile $Notes
try {
    gh release create $Tag $ExePath --repo $Repo --title $Tag --notes-file $notesFile
    Assert-Native $? "gh release create"
} finally {
    Remove-Item -LiteralPath $notesFile -Force -ErrorAction SilentlyContinue
}

Write-Host ""
Write-Host "릴리스 $Tag 완료: https://github.com/$Repo/releases/tag/$Tag" -ForegroundColor Green
