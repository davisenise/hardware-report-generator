@echo off
title Gerar Relatorio de Inventario - Davi Senise TI
color 0A

echo ================================================
echo    GERADOR DE RELATORIO DE INVENTARIO (HTML)
echo    Tecnico: Davi Senise - Suporte TI
echo ================================================
echo.
echo  Coletando dados e gerando o relatorio...
echo.

:: -ExecutionPolicy Bypass ignora a politica de execucao E o
:: Mark of the Web so nesta execucao - sem mexer na maquina.
PowerShell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Gerar-RelatorioHTML.ps1"
