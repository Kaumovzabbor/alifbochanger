@echo off
rem Alifbo Konverter: .exe yig'ish (Windows). Virtual muhit ichida ishga tushiring.
setlocal
cd /d "%~dp0"

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

pyinstaller --noconfirm --onefile --windowed --name "AlifboKonverter" --icon assets\icon.ico --add-data "app\exceptions.json;app" --add-data "app\strings_uz.json;app" --add-data "assets\icon.ico;assets" --collect-all tkinterdnd2 main.py
if errorlevel 1 (
  echo.
  echo XATO: .exe yig'ilmadi. Yuqoridagi xabarlarni tekshiring.
  exit /b 1
)

echo.
echo TAYYOR: dist\AlifboKonverter.exe
endlocal
