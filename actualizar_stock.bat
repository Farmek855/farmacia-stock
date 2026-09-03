@echo off
REM Ejecuta el script de actualizacion de stock y deja la ventana abierta
REM para ver el resultado (o los errores, si los hay).

cd /d "%~dp0"
python exportar_stock.py

echo.
echo -----------------------------------------
echo Presiona una tecla para cerrar esta ventana...
pause > nul
