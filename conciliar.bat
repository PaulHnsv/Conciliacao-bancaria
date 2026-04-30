@echo off
chcp 65001 >nul
title Conciliacao Bancaria

echo ============================================
echo    CONCILIACAO BANCARIA AUTOMATICA
echo ============================================
echo.

:: Verifica se Python esta instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Python nao encontrado.
    echo.
    echo Por favor, instale o Python 3.8+ em:
    echo   https://www.python.org/downloads/
    echo.
    echo Marque a opcao "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)

:: Vai para o diretorio do .bat
cd /d "%~dp0"

:: Instala dependencias usando wheels pre-compilados (sem necessidade de compilar)
echo Verificando dependencias...
python -m pip install -r requirements.txt --prefer-binary --quiet --disable-pip-version-check
if errorlevel 1 (
    echo AVISO: Nao foi possivel instalar dependencias automaticamente.
    echo.
    echo Tente instalar manualmente abrindo um terminal nesta pasta e executando:
    echo   pip install ofxparse pandas openpyxl pdfplumber --prefer-binary
    echo.
    pause
    exit /b 1
)

:: Executa o script principal
echo Abrindo seletor de arquivo...
echo.
python main.py

:: Se terminar com erro, mostra mensagem
if errorlevel 1 (
    echo.
    echo O programa encerrou com um erro.
    pause
)
