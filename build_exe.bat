@echo off
rem Builds dist\IED2OMICRON_<version>.exe (version from ied2omicron.py, VERSION = ...)
cd /d "%~dp0"
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller
for /f "delims=" %%n in ('python exe_name.py') do set EXENAME=%%n
echo Building %EXENAME%.exe ...
pyinstaller --clean --onefile --windowed --name %EXENAME% ied2omicron.py
echo.
echo Result: dist\%EXENAME%.exe
pause
