@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .venv (
    echo Creating virtual environment...
    python -m venv .venv || goto :error
)
call .venv\Scripts\activate.bat
echo Installing libraries...
pip install -q -r requirements.txt || goto :error
python parser.py --out books_example.xlsx || goto :error
echo.
echo DONE. Opening books_example.xlsx
start "" books_example.xlsx
pause
exit /b 0
:error
echo.
echo ERROR - take a screenshot and send it to Claude
pause
