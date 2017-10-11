@echo off
REM powershell -Command Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
powershell -ExecutionPolicy Bypass -noLogo -File speech.ps1
REM powershell -Command Set-ExecutionPolicy -Scope CurrentUser Restricted