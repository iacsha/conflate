@echo off
setlocal enabledelayedexpansion

title Conflate Build

echo.
echo  =======================================
echo    Conflate v1.0 -- EXE Build Script
echo  =======================================
echo.
echo  NOTE: Run this from plain CMD, not Anaconda Prompt.
echo  Anaconda Prompt injects paths that break the build.
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 ( echo  ERROR: python not found on PATH. & pause & exit /b 1 )
for /f "tokens=*" %%V in ('python --version') do echo  Using %%V
echo.

echo [1/5] Creating clean build environment...
if exist build_venv rmdir /s /q build_venv
python -m venv build_venv
if %errorlevel% neq 0 ( echo  ERROR: venv failed. & pause & exit /b 1 )

echo.
echo [2/5] Installing dependencies...
build_venv\Scripts\python.exe -m pip install --upgrade pip --quiet
build_venv\Scripts\pip.exe install --quiet ^
    pyinstaller ^
    customtkinter ^
    pandas ^
    rapidfuzz ^
    "scikit-learn==1.5.2" ^
    scipy ^
    openpyxl ^
    pillow ^
    pywin32-ctypes
if %errorlevel% neq 0 ( echo  ERROR: pip failed. & pause & exit /b 1 )

for /f "tokens=*" %%V in ('build_venv\Scripts\python.exe -c "import sklearn; print(sklearn.__version__)"') do (
    echo  sklearn version in venv: %%V
)
echo  Dependencies installed.

echo.
echo [3/5] Cleaning previous output...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

echo.
echo [4/5] Building...
echo  Clearing any injected Python paths...
set PYTHONPATH=
set PYTHONHOME=

build_venv\Scripts\python.exe -m PyInstaller Conflate.spec --noconfirm 2>&1

if exist "dist\Conflate.exe" (
    set BUILD_MODE=single
    goto :verify
)

echo.
echo  Single-file failed. Trying folder mode...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist
build_venv\Scripts\python.exe -m PyInstaller Conflate_folder.spec --noconfirm 2>&1

if exist "dist\Conflate\Conflate.exe" (
    set BUILD_MODE=folder
    goto :verify
)

echo  ERROR: Both build modes failed.
pause & exit /b 1

:verify
echo.
echo [5/5] Packaging output...

if "!BUILD_MODE!"=="single" (
    for %%I in ("dist\Conflate.exe") do set SIZE=%%~zI
    set /a SIZE_MB=!SIZE! / 1048576
    powershell -Command "Compress-Archive -Path 'dist\Conflate.exe' -DestinationPath 'dist\Conflate_v1.zip' -Force"
    echo.
    echo  =======================================
    echo    BUILD SUCCESSFUL  (single-file)
    echo    EXE:  dist\Conflate.exe  (!SIZE_MB! MB)
    echo    ZIP:  dist\Conflate_v1.zip
    echo  =======================================
)

if "!BUILD_MODE!"=="folder" (
    powershell -Command "Compress-Archive -Path 'dist\Conflate' -DestinationPath 'dist\Conflate_v1.zip' -Force"
    echo.
    echo  =======================================
    echo    BUILD SUCCESSFUL  (folder mode)
    echo    ZIP:  dist\Conflate_v1.zip
    echo  =======================================
)

echo.
echo  Ship dist\Conflate_v1.zip - users unzip and double-click Conflate.exe
echo.
endlocal
pause
