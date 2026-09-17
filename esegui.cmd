@echo off
REM Lancia lo script PowerShell nella STESSA cartella di questo .cmd,
REM qualunque essa sia (Download, Desktop, USB, o %USERPROFILE%\MailClient
REM dopo l'installazione). %~dp0 si risolve sempre alla cartella corrente
REM di questo file, quindi funziona prima E dopo il self-install.
powershell -ExecutionPolicy Bypass -File "%~dp0clientmail_outlook_v20.ps1"
