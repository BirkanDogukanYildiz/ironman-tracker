@echo off
REM IRONMAN Takip - telefondan da erisilebilecek sekilde baslatir (Windows)
cd /d "%~dp0"
if not exist ".venv" (
  echo Sanal ortam olusturuluyor...
  python -m venv .venv
  .venv\Scripts\python -m pip install --quiet --upgrade pip
  .venv\Scripts\pip install --quiet -r requirements.txt
  echo Kurulum tamam.
)
echo.
echo Windows guvenlik duvari sorarsa "Ozel aglar" kutusunu isaretleyip
echo ERISIME IZIN VER deyin - aksi halde telefondan baglanamazsiniz.
echo.
.venv\Scripts\python app.py --telefon %*
pause
