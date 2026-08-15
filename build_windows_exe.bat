@echo off
REM Build a single-file Windows exe of the Ernest Money Tracker.
REM Run this on a Windows PC with Python installed, from this folder.

pip install pyinstaller pandas openpyxl fpdf2 python-dateutil reportlab pypdf || exit /b 1

pyinstaller --onefile --windowed --name ErnestMoneyTracker ^
  --add-data "blank temp 1.pdf;." ^
  --add-data "fonts;fonts" ^
  "Time_and_Tune_Ernest_Money_Tracker_V8_AUTO_EMAIL_RELIABILITY_FIXED.py"

echo.
echo Done. The exe is in the dist\ folder: dist\ErnestMoneyTracker.exe
echo Copy it to any PC - no Python installation needed there.
pause
