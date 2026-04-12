@echo off
setlocal
pushd "%~dp0" || exit /b 1

set "PYW="
if exist "C:\Users\Administrator\AppData\Local\Programs\Python\Python310\pythonw.exe" set "PYW=C:\Users\Administrator\AppData\Local\Programs\Python\Python310\pythonw.exe"
if not defined PYW if exist "C:\Users\Administrator\AppData\Local\Programs\Python\Python311\pythonw.exe" set "PYW=C:\Users\Administrator\AppData\Local\Programs\Python\Python311\pythonw.exe"
if not defined PYW for %%I in (pythonw.exe) do set "PYW=%%~$PATH:I"

if not defined PYW (
  echo [ERROR] pythonw.exe not found.
  pause
  exit /b 1
)

if not exist "main.py" (
  echo [ERROR] main.py not found.
  pause
  exit /b 1
)

start "" "%PYW%" "main.py"
exit /b 0
