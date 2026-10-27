@echo off
rem Starter opslag af Send on behalf-brugere i Exchange Online (GUI)
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Find-SendOnBehalfBrugere.ps1"
