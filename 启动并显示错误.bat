@echo off
setlocal
cd /d "%~dp0"

echo [1/4] cwd: %cd%

set "PY="
if exist "C:\Users\Administrator\AppData\Local\Programs\Python\Python310\python.exe" set "PY=C:\Users\Administrator\AppData\Local\Programs\Python\Python310\python.exe"
if not defined PY if exist "C:\Users\Administrator\AppData\Local\Programs\Python\Python311\python.exe" set "PY=C:\Users\Administrator\AppData\Local\Programs\Python\Python311\python.exe"
if not defined PY for %%I in (python.exe) do set "PY=%%~$PATH:I"

if not defined PY goto :NO_PY
if not exist "%~dp0main.py" goto :NO_MAIN

echo [2/4] python: %PY%
echo [3/4] launching...
"%PY%" "%~dp0main.py"
set "EC=%ERRORLEVEL%"

echo [4/4] exit code: %EC%
if not "%EC%"=="0" echo startup failed, send this output to me.
pause
exit /b %EC%

:NO_PY
echo [ERROR] python.exe not found.
pause
exit /b 1

:NO_MAIN
echo [ERROR] main.py not found.
pause
exit /b 1
