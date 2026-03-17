@echo off
title Teams Chat Viewer - PST Server
color 0A
echo.
echo  ================================================
echo   Teams Chat Viewer - PST Conversion Server
echo  ================================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python not found!
    echo.
    echo  Please install Python from https://python.org
    echo  Make sure to check "Add Python to PATH" during install.
    echo.
    pause
    exit /b 1
)

:: Install dependencies silently
echo  Checking dependencies...
pip install libpff-python >nul 2>&1

echo  Starting server...
echo.
python server.py
pause
