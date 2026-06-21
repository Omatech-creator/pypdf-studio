@echo off
REM ===========================================================================
REM  Build a PORTABLE single-file PyPDF Studio executable with PyInstaller.
REM
REM  Output:  dist\PyPDFStudio.exe   (self-contained, no Python needed)
REM
REM  The app stores its settings/logs in a "PyPDFStudio-Data" folder next to
REM  the .exe, so you can drop the exe on a USB stick and run it anywhere.
REM ===========================================================================
setlocal
REM Use the full Python install that has all the packages (NOT the Store stub,
REM which PyInstaller cannot read due to MSIX sandboxing).
set "PYEXE=C:\Python314\python.exe"
if not exist "%PYEXE%" set "PYEXE=python"
cd /d "%~dp0"

echo Ensuring PyInstaller is installed...
"%PYEXE%" -m pip install --upgrade pyinstaller || goto :err

echo.
echo Building portable executable (this can take a few minutes)...
"%PYEXE%" -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name "PyPDFStudio" ^
  --collect-all PySide6 ^
  --collect-all fitz ^
  --collect-all pymupdf ^
  --hidden-import pypdf ^
  --hidden-import PIL ^
  --hidden-import reportlab ^
  --hidden-import pytesseract ^
  --hidden-import docx ^
  --hidden-import pptx ^
  --hidden-import openpyxl ^
  --exclude-module cv2 ^
  --exclude-module tkinter ^
  main.py || goto :err

echo.
echo ============================================================
echo  DONE.  Portable executable:  dist\PyPDFStudio.exe
echo  Copy that single file anywhere and run it.
echo ============================================================
pause
exit /b 0

:err
echo.
echo BUILD FAILED. See messages above.
pause
exit /b 1
