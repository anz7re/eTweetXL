@echo off
setlocal EnableDelayedExpansion

if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit

set X=0
set dr=C:
set env=%HOMEPATH%
set loc=\.z7\autokit\etweetxl\mtsett\webcheck.txt

for /f "tokens=10 delims=," %%A in ('tasklist /fo csv /v /fi "imagename eq firefox.exe"' ) do set /a X+=1 && set URL=%%A && if !X!==1 goto EscLoop 
exit

:EscLoop
echo %URL% > %dr%%env%%loc%

exit