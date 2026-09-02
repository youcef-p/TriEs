@echo off
if not exist "templates\index.html" (
    echo.
    echo ERROR: templates\index.html was not found in this folder.
    echo This app needs a subfolder named "templates" containing index.html,
    echo sitting right next to app.py:
    echo.
    echo   this-folder\
    echo     app.py
    echo     templates\
    echo       index.html
    echo.
    echo Put index.html inside a "templates" subfolder and run this again.
    pause
    exit /b 1
)

echo Starting... your browser will open automatically.
python app.py
