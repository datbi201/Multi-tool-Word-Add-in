@echo off
setlocal
cd /d "%~dp0"
set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
if not exist "%FW%\csc.exe" set FW=%WINDIR%\Microsoft.NET\Framework\v4.0.30319
if not exist "%FW%\csc.exe" (
  echo Khong tim thay csc.exe cua .NET Framework 4.x. Hay bat tinh nang .NET Framework 4.8 cua Windows.
  pause
  exit /b 1
)
if not exist bin mkdir bin
"%FW%\csc.exe" /nologo /target:library /platform:anycpu /codepage:65001 /optimize+ /out:bin\TroLyKyThuat.dll /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /r:System.Security.dll /r:System.Xml.dll src\*.cs
if errorlevel 1 (
  echo.
  echo BIEN DICH LOI. Hay chup man hinh cac dong loi o tren va kem theo khi bao loi.
  pause
  exit /b 1
)
echo.
echo BIEN DICH THANH CONG: bin\TroLyKyThuat.dll
endlocal
