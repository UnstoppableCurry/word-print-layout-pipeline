@echo off
set LOG=C:\printnode-setup.log
echo ==== %date% %time% setup start ==== >> %LOG% 2>&1
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup.ps1" >> %LOG% 2>&1
echo ==== %date% %time% setup exit %errorlevel% ==== >> %LOG% 2>&1
