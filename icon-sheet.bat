@echo off
setlocal
cd /d "%~dp0"
set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
if not exist "%FW%\csc.exe" set FW=%WINDIR%\Microsoft.NET\Framework\v4.0.30319
if not exist bin mkdir bin
"%FW%\csc.exe" /nologo /target:exe /platform:anycpu /codepage:65001 /out:bin\icon-sheet.exe /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll tools\IconSheet.cs src\Icons.cs
if errorlevel 1 (
  echo BIEN DICH LOI icon-sheet.
  exit /b 1
)
bin\icon-sheet.exe src\Icons.cs icons-sheet.png
endlocal
