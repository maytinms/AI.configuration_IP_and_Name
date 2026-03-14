@echo off
cd /d %~dp0
powershell -ExecutionPolicy Bypass -Command "Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -File \"%~dp0Set-NetworkConfig.ps1\"' -Verb RunAs"
