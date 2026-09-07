@echo off
REM DocConvert launcher - installs deps (first run) and starts the app.
cd /d "%~dp0"
echo Installing/updating dependencies...
python -m pip install -r requirements.txt --quiet
echo Starting DocConvert at http://127.0.0.1:5000 ...
start "" http://127.0.0.1:5000
python app.py
