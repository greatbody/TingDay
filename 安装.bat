@echo off
for /f "delims=" %%i in ('dir *.ocx *.dll /b') do (
copy "%%i" "%windir%\system32\" /y
regsvr32 /s "%windir%\system32\%%i"
)
start tingday.exe
pause 