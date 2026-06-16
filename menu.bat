@echo off
"C:\Users\Bramdon.Blanco\AppData\Local\Programs\Python\Python312\python.exe" "C:\Users\Bramdon.Blanco\Desktop\SoffgitDeck\menu.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Presiona Enter para cerrar...
    pause > nul
)
