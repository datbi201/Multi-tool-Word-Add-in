@echo off
setlocal
cd /d "%~dp0"
rem Kiem tra tren Word that (mo mot phien Word an qua COM): bang co o gop, chuan hoa doan van nhieu doan, va ten lenh idMso cua the Cong cu. Can cai Word.
set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
if not exist bin mkdir bin
"%FW%\csc.exe" /nologo /codepage:65001 /target:exe /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /r:System.Security.dll /r:System.Xml.dll /out:bin\real-word-merged.exe src\*.cs tools\RealWordMerged.cs
if errorlevel 1 (pause & exit /b 1)
bin\real-word-merged.exe
"%FW%\csc.exe" /nologo /codepage:65001 /target:exe /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /r:System.Security.dll /r:System.Xml.dll /out:bin\real-word-align.exe src\*.cs tools\RealWordAlign.cs
if errorlevel 1 (pause & exit /b 1)
bin\real-word-align.exe
"%FW%\csc.exe" /nologo /codepage:65001 /target:exe /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /r:System.Security.dll /r:System.Xml.dll /out:bin\real-word-mso.exe src\*.cs tools\RealWordMso.cs
if errorlevel 1 (pause & exit /b 1)
bin\real-word-mso.exe
pause
endlocal
