@echo off
REM ==========================================================
REM run_export.bat - Lance l'export des contrats Oracle
REM ==========================================================

REM --- A ADAPTER ---
set "USER=ERP_PPM_DM.PRISM"
set "PASS=ERP_PPM_DMPrism12345@"
set "SCRIPT=[Dossier]\test.ps1"
set "CONTRACTS=[Dossier]\contrats.txt"
set "OUTPUT=[Dossier]"
set "BASEURL=https://iabtgs-dev5.fa.ocs.oraclecloud.com"
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