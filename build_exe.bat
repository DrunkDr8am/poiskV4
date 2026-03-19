@echo off
setlocal

cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -m PyInstaller --noconfirm "guiV4.spec"
) else (
    python -m PyInstaller --noconfirm "guiV4.spec"
)

endlocal
