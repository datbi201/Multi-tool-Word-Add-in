@echo off
setlocal
net session >nul 2>&1
if errorlevel 1 (
  echo Can chay voi quyen quan tri. Bam chuot phai vao uninstall.bat, chon Run as administrator.
  pause
  exit /b 1
)
echo Dong Word truoc khi go. Bam phim bat ky de tiep tuc.
pause >nul
set DEST=%ProgramData%\TroLyKyThuat\bin
set PLAT=x64
for /f "tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v Platform 2^>nul') do set PLAT=%%a
if /i "%PLAT%"=="x86" (
  set FW=%WINDIR%\Microsoft.NET\Framework\v4.0.30319
) else (
  set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
)
if exist "%DEST%\TroLyKyThuat.dll" "%FW%\regasm.exe" /nologo /unregister "%DEST%\TroLyKyThuat.dll"
reg delete "HKCU\Software\Microsoft\Office\Word\Addins\TroLyKyThuat.Connect" /f >nul 2>&1
del /q "%DEST%\TroLyKyThuat.dll" >nul 2>&1
echo DA GO CAI DAT. Cai dat cua ban (settings.ini) van con trong %%APPDATA%%\TroLyKyThuat.
endlocal
pause
