@echo off
cd /d "%~dp0"
echo.
echo Lingua - Korean-Chinese dictionary
echo First run installs packages. Please wait.
echo A browser window opens when it is ready.
echo Close this black window to stop the app.
echo.

where py >nul 2>&1
if %errorlevel%==0 (
  py -3 -m pip install -r requirements.txt
  if errorlevel 1 goto fail
  py -3 -m streamlit run app.py
  goto end
)

where python >nul 2>&1
if %errorlevel%==0 (
  python -m pip install -r requirements.txt
  if errorlevel 1 goto fail
  python -m streamlit run app.py
  goto end
)

echo Python is not installed.
echo 1. Open https://www.python.org/downloads/
echo 2. Install Python 3.10 or newer.
echo 3. On the first install screen, check "Add python.exe to PATH".
echo 4. Double-click this file again.
pause
goto end

:fail
echo.
echo Package install failed. Check the internet connection and run this file again.
pause

:end
