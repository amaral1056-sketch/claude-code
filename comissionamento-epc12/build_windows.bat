@echo off
rem Gera dist\ComissionamentoEPC12.exe (executavel unico, nao precisa instalar).
rem Requer Python 3.10+ no computador onde o executavel e gerado (o usuario final nao precisa de Python).
cd /d "%~dp0"
python -m pip install -r requirements-dev.txt || exit /b 1
python -m pytest testes -q || exit /b 1
python -m PyInstaller --noconfirm --clean ComissionamentoEPC12.spec || exit /b 1
echo.
echo Pronto: dist\ComissionamentoEPC12.exe
