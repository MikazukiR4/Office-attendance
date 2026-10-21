@echo off
setlocal
cd /d "%~dp0"

set "VENV_PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%VENV_PYTHON%" (
    echo Run run_app.bat once to create the virtual environment and install app dependencies.
    pause
    exit /b 1
)

"%VENV_PYTHON%" -m pip install -r "%~dp0requirements-build.txt"
if errorlevel 1 (
    echo Could not install packaging dependencies.
    pause
    exit /b 1
)

"%VENV_PYTHON%" -m PyInstaller --noconfirm --clean --onefile --console ^
    --name AttendanceGenerator ^
    --collect-all streamlit ^
    --collect-all docxtpl ^
    --collect-all docx ^
    --collect-all docx2pdf ^
    --collect-all lxml ^
    --collect-all pymupdf ^
    --collect-all win32com ^
    --hidden-import pythoncom ^
    --hidden-import pywintypes ^
    --hidden-import win32timezone ^
    --add-data "app.py;." ^
    --add-data "attendance_config.py;." ^
    --add-data "generate_attendance.py;." ^
    --add-data ".streamlit\config.toml;.streamlit" ^
    --add-data "static\embedpdf.html;static" ^
    --add-data "ปกลูกจ้าง.docx;." ^
    --add-data "บัญชีลงเวลาการปฏิบัติราชการของลูกจ้าง.docx;." ^
    --add-data "ปกพนักงานราชการ.docx;." ^
    --add-data "บัญชีลงเวลา-พนักงานราช.docx;." ^
    --add-data "ปกยาม.docx;." ^
    --add-data "บัญชีลงเวลาปฏิบัติหน้าที่ยามรักษาการณ์ป.doc;." ^
    portable_launcher.py
if errorlevel 1 (
    echo The portable app build failed.
    pause
    exit /b 1
)

echo.
echo Portable app created: %~dp0dist\AttendanceGenerator.exe
echo Microsoft Word must be installed on the target computer.
pause
endlocal
