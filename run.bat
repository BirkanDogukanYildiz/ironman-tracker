@echo off
REM IRONMAN Takip - kurulum + calistirma (Windows)
cd /d "%~dp0"
if not exist ".venv" (
  echo Sanal ortam olusturuluyor...
  python -m venv .venv
  .venv\Scripts\python -m pip install --quiet --upgrade pip
  .venv\Scripts\pip install --quiet -r requirements.txt
  echo Kurulum tamam.
)
.venv\Scripts\python app.py %*
pause
