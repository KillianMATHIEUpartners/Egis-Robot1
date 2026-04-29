@echo off
REM ==========================================================
REM run_export.bat - Lance l'export des contrats Oracle
REM ==========================================================

REM --- A ADAPTER ---
set "USER=ERP_PPM_DM.PRISM"
set "PASS=Prism987654321*"
set "SCRIPT=C:\Users\KillianMathieu\Documents\UiPath\Egis\Egis-Robot1\Data\Input\ApiCall.ps1"
set "CONTRACTS=C:\Users\KillianMathieu\Documents\UiPath\Egis\Egis-Robot1\Data\Input\contrats.txt"
set "OUTPUT=C:\Users\KillianMathieu\Documents\UiPath\Egis\Egis-Robot1\Data\Output"
set "BASEURL=https://iabtgs-dev2.fa.ocs.oraclecloud.com"
set "BATCH=25"

REM --- NE PAS TOUCHER ---
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" ^
  -User "%USER%" ^
  -Password "%PASS%" ^
  -ContractsFile "%CONTRACTS%" ^
  -BaseUrl "%BASEURL%" ^
  -OutputFolder "%OUTPUT%" ^
  -BatchSize %BATCH%

pause