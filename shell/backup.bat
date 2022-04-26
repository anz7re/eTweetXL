@echo off

set dr=C:
set env=%HOMEPATH%
set loc=\.z7\autokit\etweetxl\shell\win\backup.ps1

powershell -executionpolicy remotesigned -File %dr%%env%%loc%

exit