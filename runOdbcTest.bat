@echo off
setlocal

cls
echo ***********************
echo ***** runOdbcTest *****
echo ***********************

setlocal EnableDelayedExpansion

echo Running tests as 64 bit process...
c:\Windows\system32\cscript.exe //nologo odbcTest.wsf 64

