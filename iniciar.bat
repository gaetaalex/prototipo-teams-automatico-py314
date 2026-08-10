@echo off
setlocal
cd /d "%~dp0"

echo Verificando o Python 3.14 convencional...
set "PYTHON_CMD="

py -3.14 -c "import sys,sysconfig; assert sys.version_info[:2] == (3,14); assert not bool(sysconfig.get_config_var('Py_GIL_DISABLED'))" >nul 2>&1
if not errorlevel 1 set "PYTHON_CMD=py -3.14"

if not defined PYTHON_CMD (
  python -c "import sys,sysconfig; assert sys.version_info[:2] == (3,14); assert not bool(sysconfig.get_config_var('Py_GIL_DISABLED'))" >nul 2>&1
  if not errorlevel 1 set "PYTHON_CMD=python"
)

if not defined PYTHON_CMD (
  echo ERRO: Python 3.14 convencional nao foi encontrado.
  echo Esta versao nao deve ser iniciada com Python free-threaded.
  pause
  exit /b 1
)

if exist ".venv\Scripts\python.exe" (
  .venv\Scripts\python.exe -c "import sys,sysconfig; assert sys.version_info[:2] == (3,14); assert not bool(sysconfig.get_config_var('Py_GIL_DISABLED'))" >nul 2>&1
  if errorlevel 1 (
    echo ERRO: A pasta .venv pertence a outra versao do Python.
    echo Renomeie ou remova somente a pasta .venv e execute novamente.
    pause
    exit /b 1
  )
)

if not exist ".venv\Scripts\python.exe" (
  echo Criando o ambiente virtual...
  %PYTHON_CMD% -m venv .venv
  if errorlevel 1 goto :erro_venv
)

.venv\Scripts\python.exe -c "import flask,openpyxl,msal,requests,dotenv,cryptography" >nul 2>&1
if errorlevel 1 (
  if exist "pacotes_offline\*.whl" (
    echo Instalando dependencias pelo pacote offline...
    .venv\Scripts\python.exe -m pip install --no-index --find-links="pacotes_offline" -r requirements.txt
  ) else (
    echo Instalando dependencias pelo repositorio configurado no pip...
    .venv\Scripts\python.exe -m pip install -r requirements.txt
  )
  if errorlevel 1 goto :erro_dependencias
)

if not exist ".env" copy ".env.example" ".env" >nul

echo Iniciando em http://127.0.0.1:5056
.venv\Scripts\python.exe app.py
goto :fim

:erro_venv
echo ERRO: Nao foi possivel criar o ambiente virtual.
pause
exit /b 1

:erro_dependencias
echo.
echo ERRO: Nao foi possivel instalar as dependencias.
echo Em rede corporativa, coloque os arquivos WHL em pacotes_offline
echo ou solicite a URL do repositorio Python homologado pela TI.
pause
exit /b 1

:fim
pause
endlocal
