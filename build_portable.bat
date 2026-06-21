@echo off
REM ===========================================================================
REM  Build a SLIM, PORTABLE PyPDF Studio (onedir) with PyInstaller.
REM
REM  Output:  dist\PyPDFStudio\   (a self-contained folder; no Python needed)
REM           Run dist\PyPDFStudio\PyPDFStudio.exe   -- starts instantly.
REM
REM  Strategy: bundle everything (so Qt plugins are found), then delete the
REM  large Qt modules the app never uses (WebEngine, Quick/QML, 3D, Charts,
REM  Multimedia, ...). This cuts the build from ~760 MB to ~245 MB.
REM
REM  Settings/logs are written to a "PyPDFStudio-Data" folder next to the exe,
REM  so the whole folder is portable (copy it to a USB stick and run anywhere).
REM ===========================================================================
setlocal
set "PYEXE=C:\Python314\python.exe"
if not exist "%PYEXE%" set "PYEXE=python"
cd /d "%~dp0"

echo Ensuring PyInstaller is installed...
"%PYEXE%" -m pip install --upgrade pyinstaller || goto :err

echo.
echo Building portable app (this can take a few minutes)...
"%PYEXE%" -m PyInstaller --noconfirm --clean --onedir --windowed ^
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
echo Pruning unused Qt modules to slim the bundle...
set "QT=dist\PyPDFStudio\_internal\PySide6"
for %%M in (
  Qt6WebEngine QtWebEngine Qt6WebChannel QtWebChannel Qt6WebSockets QtWebSockets Qt6WebView QtWebView
  Qt6Quick QtQuick Qt6Qml QtQml Qt63D Qt3D
  Qt6Charts QtCharts Qt6DataVisualization QtDataVisualization Qt6Graphs QtGraphs
  Qt6Multimedia QtMultimedia Qt6SpatialAudio QtSpatialAudio
  Qt6Designer QtDesigner Qt6Help QtHelp Qt6UiTools QtUiTools
  Qt6Sql QtSql Qt6Test QtTest Qt6Bluetooth QtBluetooth Qt6Nfc QtNfc Qt6Serial QtSerial
  Qt6Positioning QtPositioning Qt6Sensors QtSensors Qt6Location QtLocation
  Qt6RemoteObjects QtRemoteObjects Qt6Scxml QtScxml Qt6StateMachine QtStateMachine
  Qt6TextToSpeech QtTextToSpeech Qt6Pdf QtPdf Qt6Concurrent QtConcurrent
  avcodec avformat avutil swscale swresample
) do del /q "%QT%\%%M*" >nul 2>&1
for %%D in (resources qml translations) do rmdir /s /q "%QT%\%%D" >nul 2>&1

echo.
echo ============================================================
echo  DONE.  Portable app folder:  dist\PyPDFStudio
echo  Run:   dist\PyPDFStudio\PyPDFStudio.exe
echo  Copy the whole PyPDFStudio folder anywhere to run it.
echo ============================================================
pause
exit /b 0

:err
echo.
echo BUILD FAILED. See messages above.
pause
exit /b 1
