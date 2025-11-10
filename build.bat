@echo off
echo ========================================
echo File Checker - Build Script
echo ========================================
echo.

echo [1/4] Installing dependencies...
python -m pip install -r requirements.txt
python -m pip install pyinstaller

echo.
echo [2/4] Creating necessary directories...
if not exist "templates" mkdir templates
if not exist "uploads" mkdir uploads
if not exist "reports" mkdir reports

echo.
echo [3/4] Building executable with PyInstaller...
python -m PyInstaller --clean FileChecker.spec

echo.
echo [4/4] Build complete!
echo.
echo Your executable is located in: dist\FileChecker.exe
echo.
echo ========================================
echo Build finished successfully!
echo ========================================
pause