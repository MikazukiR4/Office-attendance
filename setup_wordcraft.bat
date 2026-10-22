@echo off
setlocal
cd /d "%~dp0"

set "WORDCRAFT_VERSION=0.9.0"
set "WORDCRAFT_DIR=%~dp0.tools\wordcraft"
set "WORDCRAFT_ARCHIVE=%TEMP%\wordcraft-%WORDCRAFT_VERSION%-windows-x64-portable.zip"
set "WORDCRAFT_URL=https://github.com/storytold/wordcraft/releases/download/v%WORDCRAFT_VERSION%/wordcraft-%WORDCRAFT_VERSION%-windows-x64-portable.zip"

if exist "%WORDCRAFT_DIR%\wordcraft-cli.exe" exit /b 0
if not exist "%WORDCRAFT_DIR%" mkdir "%WORDCRAFT_DIR%"

echo Downloading WordCraft %WORDCRAFT_VERSION%...
curl.exe -fL --retry 3 -o "%WORDCRAFT_ARCHIVE%" "%WORDCRAFT_URL%"
if errorlevel 1 (
    echo Could not download WordCraft.
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -LiteralPath '%WORDCRAFT_ARCHIVE%' -DestinationPath '%WORDCRAFT_DIR%' -Force"
if errorlevel 1 (
    echo Could not extract WordCraft.
    exit /b 1
)

for /r "%WORDCRAFT_DIR%" %%F in (wordcraft-cli.exe) do (
    if not "%%~dpF"=="%WORDCRAFT_DIR%\" copy /y "%%F" "%WORDCRAFT_DIR%\wordcraft-cli.exe" >nul
)

if not exist "%WORDCRAFT_DIR%\wordcraft-cli.exe" (
    echo WordCraft CLI was not found in the downloaded archive.
    exit /b 1
)

del /q "%WORDCRAFT_ARCHIVE%" >nul 2>nul
echo WordCraft is ready.
exit /b 0
