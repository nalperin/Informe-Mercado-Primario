@echo off
REM ============================================================
REM  Generador automatico — Informe Mercado Primario de ONs
REM ============================================================

cd /d "C:\Users\nalperin\OneDrive - One618 Financial Services S.A.U\Escritorio\Claude\Informe ONs"

"C:\Users\nalperin\AppData\Local\Python\pythoncore-3.14-64\python.exe" Logica/main.py >> Logica/log_generacion.txt 2>&1

echo [%date% %time%] Informe generado >> Logica/log_generacion.txt

REM Abrir el informe actualizado en el browser
start "" "Output\informe_ultimo.html"
