@echo off
setlocal

:: Set the virtual environment name
set VENV_DIR=venv

:: Check if the virtual environment exists, if not, create it
if not exist %VENV_DIR% (
    echo Creating virtual environment...
    python -m venv %VENV_DIR%
)

:: Activate the virtual environment
call %VENV_DIR%\Scripts\activate

:: Ensure pip is up to date before installing anything
echo Updating pip...
python -m pip install --upgrade pip
if %errorlevel% neq 0 (
    echo Failed to update pip. Press any key to exit...
    pause
    exit /b %errorlevel%
)

:: Install required packages
echo Installing dependencies...
pip install python-pptx moviepy==1.0.3 pillow lxml
if %errorlevel% neq 0 (
    echo Failed to install dependencies. Press any key to exit...
    pause
    exit /b %errorlevel%
)

:: Run the Python script
echo Running script...
python worshipSlideGenerator.pyw
if %errorlevel% neq 0 (
    echo Script execution encountered an error. Press any key to exit...
    pause
    exit /b %errorlevel%
)

:: Deactivate the virtual environment
deactivate

echo Script execution completed successfully.
pause
