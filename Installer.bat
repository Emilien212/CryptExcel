@echo off
if not "%1"=="am_admin" (powershell start -verb runas '%0' am_admin & exit /b)
cd %TEMP%
python --version | find "Python 3" >nul 2>nul
if %errorlevel% == 1 (
    echo Python 3 isn't install, proceed to install it
    curl https://www.python.org/ftp/python/3.8.3/python-3.8.3.exe -O python.exe
    python-3.8.3.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
    del /f python-3.8.3.exe
) else (
    echo Python 3 is already installed
)
echo Installing libraries 
pip install Twisted
pip install python-binance
pip install openpyxl
pip install datetime
pause

