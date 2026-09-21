@echo off
REM Local build fallback. The recommended way to build is GitHub Actions (see README.txt).
REM Requires Python 3.10+ installed on THIS machine, for the build step only.
echo Setting up build environment...
python -m venv build_venv
if errorlevel 1 goto :fail
call build_venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
if errorlevel 1 goto :fail
echo Building the app with PyInstaller...
pyinstaller app.spec --noconfirm
if errorlevel 1 goto :fail
echo.
echo Done. Your app is in the folder:  dist\TestLinkWEFormatter\
echo Run  dist\TestLinkWEFormatter\TestLinkWEFormatter.exe
echo Give people the WHOLE dist\TestLinkWEFormatter folder, not just the .exe.
pause
exit /b 0
:fail
echo.
echo Build failed - see the messages above.
pause
exit /b 1
