@echo off
color 0B
title Servidor Streamlit - PMO
cd /d C:\Temp\port
echo Iniciando o painel de portfólio no navegador...
python -m streamlit run portfolio.py
pause