@echo off
rem Antivirus soxta ogohlantirish bersa: --onedir variant (bitta fayl o'rniga papka).
setlocal
cd /d "%~dp0"

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

pyinstaller --noconfirm --onedir --windowed --name "AlifboKonverter" --icon assets\icon.ico --add-data "app\exceptions.json;app" --add-data "app\strings_uz.json;app" --add-data "assets\icon.ico;assets" --collect-all tkinterdnd2 main.py
if errorlevel 1 (
  echo XATO: yig'ilmadi.
  exit /b 1
)

echo.
echo TAYYOR: dist\AlifboKonverter\AlifboKonverter.exe
endlocal
