@echo off
title Silvina Editorial Assistant v0.8
cd /d "C:\00 PY\Course\silvina-editorial\silvina_editorial_v08"
echo.
echo =====================================
echo   SILVINA EDITORIAL ASSISTANT v0.8
echo =====================================
echo.
echo Iniciando Silvina... Por favor espere.
echo Chrome se abrira automaticamente.
echo.
echo NO CIERRE ESTA VENTANA mientras usa Silvina.
echo Para cerrar Silvina: presione Ctrl+C aqui.
echo.
call ..\venv312\Scripts/activate.bat
python gradio_app.py
pause