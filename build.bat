@echo off
REM Build script for MPP to XLS converter executable
REM Generates mpp-to-xls.exe using PyInstaller

echo Cleaning up previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /q *.spec

echo.
echo Building executable...
uv run pyinstaller --onefile --name "mpp-to-xls" --console --hidden-import=jpype --hidden-import=xlsxwriter mpp_to_xls_converter.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ✓ Build completed successfully!
    echo.
    echo Executable location: dist\mpp-to-xls.exe
    echo.
    echo To test the executable, run:
    echo   .\dist\mpp-to-xls.exe lib\mpxj\junit\data\SubprojectA-9.mpp output.xlsx
) else (
    echo.
    echo ✗ Build failed!
    echo.
    pause
    exit /b 1
)
