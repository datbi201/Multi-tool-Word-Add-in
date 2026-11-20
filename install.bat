@echo off
setlocal
cd /d "%~dp0"
net session >nul 2>&1
if errorlevel 1 (
  echo Can chay voi quyen quan tri. Bam chuot phai vao install.bat, chon Run as administrator.
  pause
  exit /b 1
)
echo Dong Word truoc khi cai. Bam phim bat ky de tiep tuc.
pause >nul
call build.bat
if errorlevel 1 exit /b 1

set PLAT=x64
for /f "tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v Platform 2^>nul') do set PLAT=%%a
if /i "%PLAT%"=="x86" (set FW=%WINDIR%\Microsoft.NET\Framework\v4.0.30319) else (set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319)
if exist "%ProgramData%\TroLyKyThuat\bin\TroLyKyThuat.dll" "%FW%\regasm.exe" /nologo /unregister "%ProgramData%\TroLyKyThuat\bin\TroLyKyThuat.dll" >nul 2>&1
reg delete "HKCU\Software\Microsoft\Office\Word\Addins\TroLyKyThuat.Connect" /f >nul 2>&1
set DEST=%ProgramData%\TroLyKyThuat\bin
if not exist "%DEST%" mkdir "%DEST%"
copy /y bin\TroLyKyThuat.dll "%DEST%\TroLyKyThuat.dll" >nul
if errorlevel 1 (
  echo Khong chep duoc tep vao %DEST%. Hay chac chan Word da dong.
  pause
  exit /b 1
)

set PLAT=x64
for /f "tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v Platform 2^>nul') do set PLAT=%%a
echo Office: %PLAT%
if /i "%PLAT%"=="x86" (
  set FW=%WINDIR%\Microsoft.NET\Framework\v4.0.30319
) else (
  set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
)
"%FW%\regasm.exe" /nologo /codebase "%DEST%\TroLyKyThuat.dll"
if errorlevel 1 (
  echo DANG KY COM LOI. Hay chup man hinh kem theo khi bao loi.
  pause
  exit /b 1
)

reg delete "HKCU\Software\Microsoft\Office\16.0\Word\Resiliency" /f >nul 2>&1
set KEY=HKCU\Software\Microsoft\Office\Word\Addins\TroLyKyThuat.Connect
reg add "%KEY%" /v FriendlyName /t REG_SZ /d "Tro ly Ky thuat" /f >nul
reg add "%KEY%" /v Description /t REG_SZ /d "Cong cu bang, cong thuc, thuat ngu cho bao cao ky thuat" /f >nul
reg add "%KEY%" /v LoadBehavior /t REG_DWORD /d 3 /f >nul
echo.
echo CAI DAT XONG. Mo Word, tim the "Tro ly Ky thuat" tren Ribbon.
echo Neu khong thay the, vao File - Options - Add-ins - Manage: Disabled Items - Go, bat lai add-in.
endlocal
pause
