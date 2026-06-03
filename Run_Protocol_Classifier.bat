@echo off
setlocal

REM =========================================================
REM USER CONFIGURATION
REM =========================================================

set "REPO_URL=https://github.com/DavideTomei/AI-project.git"
set "SCRIPT_NAME=protocol_classifier.py"

REM Local folder where the GitHub code will be stored
set "APP_DIR=%USERPROFILE%\ProtocolClassifierApp"

REM =========================================================
REM CHECK PYTHON
REM =========================================================

where python >nul 2>nul

if errorlevel 1 (
    echo Python was not found.
    echo Please install Python and make sure "Add Python to PATH" is enabled.
    pause
    exit /b 1
)

REM =========================================================
REM CHECK GIT
REM =========================================================

where git >nul 2>nul

if errorlevel 1 (
    echo Git was not found.
    echo Please install Git for Windows first.
    pause
    exit /b 1
)

REM =========================================================
REM DOWNLOAD OR UPDATE CODE FROM GITHUB
REM =========================================================

if not exist "%APP_DIR%" (
    echo Cloning repository from GitHub...
    git clone "%REPO_URL%" "%APP_DIR%"

    if errorlevel 1 (
        echo Failed to clone repository.
        pause
        exit /b 1
    )
) else (
    echo Updating repository from GitHub...
    cd /d "%APP_DIR%"
    git pull
)

REM =========================================================
REM GO TO APP FOLDER
REM =========================================================

cd /d "%APP_DIR%"

REM =========================================================
REM CREATE VIRTUAL ENVIRONMENT IF NEEDED
REM =========================================================

if not exist ".venv" (
    echo Creating Python virtual environment...
    python -m venv .venv
)

REM =========================================================
REM ACTIVATE VIRTUAL ENVIRONMENT
REM =========================================================

call ".venv\Scripts\activate.bat"

REM =========================================================
REM INSTALL REQUIRED PACKAGES
REM =========================================================

echo Installing required Python packages...

python -m pip install --upgrade pip
python -m pip install pymupdf requests json-repair XlsxWriter

REM =========================================================
REM RUN THE APPLICATION
REM =========================================================

echo Starting Protocol Classifier...
python "%SCRIPT_NAME%"

if errorlevel 1 (
    echo.
    echo The application stopped with an error.
    pause
    exit /b 1
)

endlocal