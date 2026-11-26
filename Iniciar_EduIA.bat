@echo off
chcp 65001 >nul
title EduIA Analytics
cd /d "%~dp0"
set "PORTA=8765"

echo ============================================
echo   EduIA Analytics - iniciando o dashboard
echo ============================================
echo.

rem ---- 1. Já está rodando? Só abre o navegador.
powershell -NoProfile -Command "try{(New-Object Net.Sockets.TcpClient('localhost',%PORTA%)).Close();exit 0}catch{exit 1}" >nul 2>&1
if %errorlevel%==0 (
    echo O dashboard ja esta aberto. Abrindo o navegador...
    start "" http://localhost:%PORTA%
    timeout /t 3 >nul
    exit /b 0
)

rem ---- 2. Encontrar o Python (3.11 ou mais novo)
set "PY="
py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)" >nul 2>&1 && set "PY=py -3"
if not defined PY (
    python -c "import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)" >nul 2>&1 && set "PY=python"
)
if not defined PY (
    echo Python 3.11 ou mais novo nao foi encontrado.
    echo.
    choice /c SN /m "Deseja instalar o Python agora (precisa de internet)"
    if errorlevel 2 goto sem_python
    winget install --id Python.Python.3.13 -e --silent --accept-source-agreements --accept-package-agreements
    echo.
    echo Python instalado. Feche esta janela e de dois cliques no atalho de novo.
    pause
    exit /b 0
)

rem ---- 3. Ambiente próprio do projeto (.venv) com as dependências
if not exist ".venv\Scripts\python.exe" (
    echo Primeira execucao: preparando o ambiente. Isso leva alguns minutos...
    %PY% -m venv .venv
    if errorlevel 1 goto erro_instalacao
)
fc /b requirements.txt ".venv\requirements.instalado" >nul 2>&1
if errorlevel 1 (
    echo Instalando as dependencias...
    ".venv\Scripts\python.exe" -m pip install --upgrade pip --quiet
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt --quiet
    if errorlevel 1 goto erro_instalacao
    copy /y requirements.txt ".venv\requirements.instalado" >nul
)

rem ---- 4. Abrir o navegador assim que o servidor subir
start "" /min powershell -NoProfile -WindowStyle Hidden -Command "for($i=0;$i -lt 90;$i++){try{(New-Object Net.Sockets.TcpClient('localhost',%PORTA%)).Close();Start-Process 'http://localhost:%PORTA%';break}catch{Start-Sleep 1}}"

rem ---- 5. Rodar o dashboard; se cair, reinicia sozinho
:rodar
echo.
echo Dashboard rodando em http://localhost:%PORTA%
echo Para encerrar, feche esta janela.
echo.
".venv\Scripts\python.exe" -m streamlit run app.py --server.port %PORTA% --server.headless true --browser.gatherUsageStats false
echo.
echo O dashboard parou. Reiniciando em 3 segundos...
timeout /t 3 >nul
goto rodar

:sem_python
echo.
echo Instale o Python em https://www.python.org/downloads/ (marque "Add python.exe to PATH")
echo e depois de dois cliques neste arquivo novamente.
pause
exit /b 1

:erro_instalacao
echo.
echo Nao foi possivel instalar as dependencias. Verifique a conexao com a internet
echo e tente de novo. Se o erro continuar, apague a pasta .venv e execute outra vez.
pause
exit /b 1
