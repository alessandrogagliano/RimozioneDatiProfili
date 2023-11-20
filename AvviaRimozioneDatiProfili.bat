@ECHO OFF
setlocal enabledelayedexpansion 

set SCRIPT_DIR=%~dp0

PowerShell -Command "& {Start-Process PowerShell -ArgumentList '-ExecutionPolicy Bypass -File ""%SCRIPT_DIR%RimozioneDatiProfili.ps1""' -Verb RunAs}"
