@echo off
echo Installing QA Screenshot Tool dependencies...
echo.
pip install --timeout 120 -r requirements.txt
echo.
echo Done. You can now run the tool with run.bat
pause
