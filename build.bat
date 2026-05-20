@echo off
:: =====================================================================
:: PYQT / PYSIDE6 PYINSTALLER COMPILATION BATCH ENGINE
:: =====================================================================
TITLE PySide6 Application Executable Compiler

:: Define Configuration Variables
SET SCRIPT_NAME=main.py
SET APP_NAME="CMI-Data Processor"
SET OUTPUT_EXE_NAME="CMI_Data_Processor"

echo ===================================================
echo 🚀 STARTING PACKAGING WORKFLOW FOR: %APP_NAME%
echo ===================================================
echo.

:: Clean historical caching maps if they exist to prevent build contamination
if exist "build" (
    echo 🧹 Cleaning historical temporary build artifacts...
    rmdir /s /q "build"
)
if exist "dist" (
    echo 🧹 Cleaning historical target distribution builds...
    rmdir /s /q "dist"
)

echo.
echo 📦 Compiling Python modules via PyInstaller Engine...
echo 💡 This may take a minute depending on your processor...
echo.

:: Run PyInstaller Compilation Routine
:: Flags Breakdown:
:: --noconfirm        Overrides files automatically without asking
:: --onedir           Bundles into an application folder (Highly recommended for PySide6)
:: --windowed         Hides the diagnostic black console window behind the GUI interface
:: --name             Specifies the target binary filename
:: --clean            Forces clear-down of PyInstaller structural system caches
pyinstaller --noconfirm --onedir --windowed --name=%OUTPUT_EXE_NAME% --clean "%SCRIPT_NAME%"

:: Verification Check Block
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ===================================================
    echo ✅ SUCCESS: Build Sequence Concluded Without Error!
    echo 📁 Navigate to the "dist\%OUTPUT_EXE_NAME%" folder.
    echo ===================================================
) else (
    echo.
    echo ===================================================
    echo ❌ ERROR: PyInstaller compilation failed.
    echo 💡 Verify that your script has no broken absolute runtime paths.
    echo ===================================================
)

pause