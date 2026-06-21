@echo off
REM ===========================================================================
REM  Launcher for PyPDF Studio.
REM  Pins the interpreter that has the required packages (C:\Python314), so
REM  double-clicking this file always uses the correct Python — NOT the
REM  Microsoft Store stub or a different 3.14 install.
REM ===========================================================================
setlocal
set "PYEXE=C:\Python314\python.exe"
if not exist "%PYEXE%" set "PYEXE=python"
cd /d "%~dp0"
"%PYEXE%" main.py
if errorlevel 1 pause
