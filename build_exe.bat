@echo off
setlocal
cd /d "%~dp0"

set "APP_NAME=TodoList"
set "OUT_EXE=dist\%APP_NAME%.exe"
set "HOLIDAY_DB=DB_holiday.xlsx"

echo ============================================
echo  Todo List - Build Executable (EXE)
echo ============================================
echo.

echo [1/3] Checking Python launcher...
py -V >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python launcher "py" was not found.
    echo         Install Python or add it to PATH, then try again.
    goto fail
)

echo.
echo [2/3] Checking dependencies...
py -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo     Installing PyInstaller...
    py -m pip install pyinstaller
    if errorlevel 1 goto fail
)

py -m pip show openpyxl >nul 2>&1
if errorlevel 1 (
    echo     Installing openpyxl for holiday DB...
    py -m pip install openpyxl
    if errorlevel 1 goto fail
)

echo.
echo [3/3] Building... takes 1-2 minutes
if exist "%OUT_EXE%" (
    del /q "%OUT_EXE%" >nul 2>&1
    if exist "%OUT_EXE%" (
        echo [ERROR] Could not replace %OUT_EXE%.
        echo         Close every running TodoList window, then run this file again.
        goto fail
    )
)

py -m PyInstaller --noconfirm --onefile --windowed --name "%APP_NAME%" --icon "assets\todolist.ico" --add-data "assets\todolist.ico;assets" --hidden-import openpyxl todo_app.py
if errorlevel 1 goto fail
if not exist "%OUT_EXE%" goto fail

if exist "%HOLIDAY_DB%" (
    copy /Y "%HOLIDAY_DB%" "dist\%HOLIDAY_DB%" >nul
    if errorlevel 1 goto fail
) else (
    echo [WARN] %HOLIDAY_DB% was not found. Built-in holidays will be used.
)

echo.
echo ============================================
echo  Done!  %OUT_EXE%
echo  Keep %HOLIDAY_DB% in the same folder as the exe.
echo  Pin TodoList.exe to your desktop / taskbar.
echo ============================================
goto end

:fail
echo.
echo ============================================
echo  Build failed. Check the message above.
echo ============================================

:end
echo.
pause
endlocal