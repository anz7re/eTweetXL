@echo off

if not DEFINED IS_MINIMIZED set IS_MINIMIZED=1 && start "" /min "%~dpnx0" %* && exit

set dr=C:
set env=%HOMEPATH%
set loc=\.z7\autokit\etweetxl\shell\win\load_login.ps1

powershell -executionpolicy remotesigned -File %dr%%env%%loc%

exit