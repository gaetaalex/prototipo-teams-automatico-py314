@echo off
setlocal
cd /d "%~dp0"

if not exist "pacotes_offline" mkdir "pacotes_offline"

echo Baixando pacotes para CPython 3.14 convencional - Windows 64 bits...
python -m pip download ^
  --dest "pacotes_offline" ^
  --only-binary=:all: ^
  --platform win_amd64 ^
  --implementation cp ^
  --python-version 3.14 ^
  --abi cp314 ^
  --abi abi3 ^
  -r requirements.txt

if errorlevel 1 (
  echo ERRO: O download nao foi concluido.
  exit /b 1
)

echo Pacote offline criado em pacotes_offline.
endlocal
