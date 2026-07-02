@echo off
setlocal

cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -m PyInstaller --noconfirm "ZSearch.spec"
) else (
    py -3 -m PyInstaller --noconfirm "ZSearch.spec"
)

endlocal
