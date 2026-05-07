@echo off
setlocal EnableDelayedExpansion

set "SRC=\\srvfile05\COMMUN\PDA"
set "DST=\\srvad05\PDA"

echo Synchronisation de %SRC% vers %DST%
robocopy "%SRC%" "%DST%" /MIR /Z /W:5 /R:3 /TEE /LOG:"%~dp0robocopy_log.txt"

echo Copie terminer.
pause
