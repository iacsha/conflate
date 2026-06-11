@echo off
setlocal enabledelayedexpansion
title Conflate Release

REM ============================================================
REM  Conflate one-command release
REM  Usage:  release.bat 1.1
REM  Requires: GitHub CLI (gh) installed and authenticated.
REM  Steps: build -> zip -> tag -> push -> publish GitHub release.
REM  Run from plain CMD in the project folder.
REM ============================================================

if "%~1"=="" (
    echo  Usage: release.bat ^<version^>     e.g.  release.bat 1.1
    pause & exit /b 1
)
set VER=%~1

echo.
echo  =======================================
echo    Conflate Release  -  v%VER%
echo  =======================================
echo.

REM --- GitHub CLI present? ---
where gh >nul 2>&1
if %errorlevel% neq 0 (
    echo  ERROR: GitHub CLI ^(gh^) not found. Install from https://cli.github.com
    pause & exit /b 1
)

REM --- gh authenticated? ---
gh auth status >nul 2>&1
if %errorlevel% neq 0 (
    echo  ERROR: gh is not authenticated. Run:  gh auth login
    pause & exit /b 1
)

REM --- tag must not already exist ---
git rev-parse -q --verify "refs/tags/v%VER%" >nul 2>&1
if %errorlevel% equ 0 (
    echo  ERROR: tag v%VER% already exists.
    pause & exit /b 1
)

REM --- working tree must be clean ---
git diff --quiet
if %errorlevel% neq 0 (
    echo  ERROR: you have uncommitted changes. Commit them before releasing.
    pause & exit /b 1
)

echo  Reminder: VERSION in Conflate.py and version_info.txt should read %VER%,
echo  and CHANGELOG.md should have a "## v%VER%" section.
echo  Press any key to continue, or Ctrl+C to abort.
pause >nul

echo.
echo [1/5] Building executable...
call build.bat
if %errorlevel% neq 0 ( echo  ERROR: build failed. & pause & exit /b 1 )

echo.
echo [2/5] Packaging dist\Conflate_v%VER%.zip ...
if exist "dist\Conflate.exe" (
    powershell -NoProfile -Command "Compress-Archive -Path 'dist\Conflate.exe' -DestinationPath 'dist\Conflate_v%VER%.zip' -Force"
) else (
    if exist "dist\Conflate\Conflate.exe" (
        powershell -NoProfile -Command "Compress-Archive -Path 'dist\Conflate' -DestinationPath 'dist\Conflate_v%VER%.zip' -Force"
    ) else (
        echo  ERROR: no build output found in dist. & pause & exit /b 1
    )
)

echo.
echo [3/5] Extracting release notes for v%VER% from CHANGELOG.md ...
powershell -NoProfile -Command "$v='%VER%'; $t=Get-Content -Raw CHANGELOG.md; $m=[regex]::Match($t,'(?ms)^##\s*v'+[regex]::Escape($v)+'\b.*?(?=^##\s*v|\Z)'); if($m.Success){[IO.File]::WriteAllText('dist\_relnotes.txt',$m.Value.Trim())}else{[IO.File]::WriteAllText('dist\_relnotes.txt','Conflate v'+$v)}"

echo.
echo [4/5] Tagging and pushing...
git tag -a v%VER% -m "Conflate v%VER%"
if %errorlevel% neq 0 ( echo  ERROR: tagging failed. & pause & exit /b 1 )
git push
if %errorlevel% neq 0 ( echo  ERROR: push failed. & pause & exit /b 1 )
git push origin v%VER%
if %errorlevel% neq 0 ( echo  ERROR: tag push failed. & pause & exit /b 1 )

echo.
echo [5/5] Creating GitHub release...
gh release create v%VER% "dist\Conflate_v%VER%.zip" --title "Conflate v%VER%" --notes-file "dist\_relnotes.txt"
if %errorlevel% neq 0 ( echo  ERROR: gh release failed. & pause & exit /b 1 )

del "dist\_relnotes.txt" >nul 2>&1

echo.
echo  =======================================
echo    RELEASE PUBLISHED  -  v%VER%
echo    https://github.com/iacsha/conflate/releases/tag/v%VER%
echo  =======================================
echo.
endlocal
pause
