@echo off
setlocal
cd /d "%~dp0"

set "VENV_PYTHON=%~dp0.venv\Scripts\python.exe"

if not exist "%VENV_PYTHON%" (
    where py >nul 2>nul
    if not errorlevel 1 (
        py -3 -m venv "%~dp0.venv"
    ) else (
        where python >nul 2>nul
        if errorlevel 1 (
            echo Python 3 is not installed. Install Python 3 and try again.
            pause
            exit /b 1
        )
        python -m venv "%~dp0.venv"
    )
    if errorlevel 1 (
        echo Could not create the virtual environment.
        pause
        exit /b 1
    )
)

if not exist "%~dp0.venv\attendance-dependencies-installed" (
    "%VENV_PYTHON%" -m pip install -r "%~dp0requirements.txt"
    if errorlevel 1 (
        echo Could not install app dependencies.
        pause
        exit /b 1
    )
    type nul > "%~dp0.venv\attendance-dependencies-installed"
)

powershell -NoProfile -Command "if ((Get-Item -LiteralPath '%~dp0requirements.txt').LastWriteTime -gt (Get-Item -LiteralPath '%~dp0.venv\attendance-dependencies-installed').LastWriteTime) { exit 1 } else { exit 0 }"
if errorlevel 1 (
    "%VENV_PYTHON%" -m pip install -r "%~dp0requirements.txt"
    if errorlevel 1 (
        echo Could not update app dependencies.
        pause
        exit /b 1
    )
    type nul > "%~dp0.venv\attendance-dependencies-installed"
)

if not exist "%~dp0.tools\wordcraft\wordcraft-cli.exe" (
    call "%~dp0setup_wordcraft.bat"
    if errorlevel 1 (
        echo WordCraft is required to generate DOCX and PDF files.
        pause
        exit /b 1
    )
)
set "WORDCRAFT_CLI=%~dp0.tools\wordcraft\wordcraft-cli.exe"
set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"
echo. | "%VENV_PYTHON%" -m streamlit run "%~dp0app.py"
if errorlevel 1 (
    echo The app could not start. Check the error above.
    pause
    exit /b 1
)

endlocal
