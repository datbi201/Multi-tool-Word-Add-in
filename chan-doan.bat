@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set OUT=%~dp0chan-doan.txt
echo === CHAN DOAN TRO LY KY THUAT === > "%OUT%"
echo Thoi gian: %DATE% %TIME% >> "%OUT%"
echo Thu muc: %~dp0 >> "%OUT%"
echo. >> "%OUT%"

echo --- Ban Office (ClickToRun) --- >> "%OUT%"
reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v Platform >> "%OUT%" 2>&1
reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v VersionToReport >> "%OUT%" 2>&1
reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v ProductReleaseIds >> "%OUT%" 2>&1
echo --- WINWORD.EXE --- >> "%OUT%"
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE" /ve >> "%OUT%" 2>&1
echo --- .NET Framework 4.x (Release) --- >> "%OUT%"
reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release >> "%OUT%" 2>&1
echo. >> "%OUT%"

echo --- Add-in cua ta trong HKCU --- >> "%OUT%"
reg query "HKCU\Software\Microsoft\Office\Word\Addins\TroLyKyThuat.Connect" /s >> "%OUT%" 2>&1
echo --- Dang ky COM (khung nhin 64-bit) --- >> "%OUT%"
reg query "HKCR\CLSID\{6F1C2B54-9A3E-4D0B-8B7A-3C5D2E1F0A47}" /s /reg:64 >> "%OUT%" 2>&1
echo --- Dang ky COM (khung nhin 32-bit) --- >> "%OUT%"
reg query "HKCR\CLSID\{6F1C2B54-9A3E-4D0B-8B7A-3C5D2E1F0A47}" /s /reg:32 >> "%OUT%" 2>&1
echo --- Danh sach add-in bi Word vo hieu hoa / gay loi (Resiliency) --- >> "%OUT%"
reg query "HKCU\Software\Microsoft\Office\16.0\Word\Resiliency" /s >> "%OUT%" 2>&1
echo. >> "%OUT%"

echo --- Cac add-in khac cua Word (HKCU, HKLM 64, HKLM 32) --- >> "%OUT%"
reg query "HKCU\Software\Microsoft\Office\Word\Addins" >> "%OUT%" 2>&1
reg query "HKLM\SOFTWARE\Microsoft\Office\Word\Addins" >> "%OUT%" 2>&1
reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\Word\Addins" >> "%OUT%" 2>&1
echo. >> "%OUT%"

echo --- Tep DLL --- >> "%OUT%"
dir "%ProgramData%\TroLyKyThuat\bin" >> "%OUT%" 2>&1
echo. >> "%OUT%"

echo --- Nhat ky cua add-in (TEMP) --- >> "%OUT%"
if exist "%TEMP%\TroLyKyThuat\log.txt" (type "%TEMP%\TroLyKyThuat\log.txt" >> "%OUT%") else (echo KHONG CO tep log trong TEMP, nghia la ma cua add-in chua chay. >> "%OUT%")
echo --- Nhat ky cua add-in (APPDATA) --- >> "%OUT%"
if exist "%APPDATA%\TroLyKyThuat\log.txt" (type "%APPDATA%\TroLyKyThuat\log.txt" >> "%OUT%") else (echo KHONG CO tep log trong APPDATA. >> "%OUT%")
echo. >> "%OUT%"

echo --- Su kien loi gan day cua Windows (Application Error, .NET Runtime, Application Hang), toi da 12 muc --- >> "%OUT%"
wevtutil qe Application /c:12 /rd:true /f:text /q:"*[System[Provider[@Name='Application Error' or @Name='.NET Runtime' or @Name='Application Hang']]]" >> "%OUT%" 2>&1
echo. >> "%OUT%"
echo --- Het --- >> "%OUT%"
echo.
echo DA GHI XONG tep chan-doan.txt trong thu muc nay. Hay bao Claude "da chay chan doan".
endlocal
pause
