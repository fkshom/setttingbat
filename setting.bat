@echo off
@REM @powershell -NoProfile -ExecutionPolicy Unrestricted "&([ScriptBlock]::Create((cat -encoding utf8 \"%~f0\" | ? {$_.ReadCount -gt 2}) -join \"`n\"))" %*
@REM @exit /b


powershell -NoProfile -ExecutionPolicy Unrestricted .\setting.ps1

pause