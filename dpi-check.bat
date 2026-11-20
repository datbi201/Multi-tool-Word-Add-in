@echo off
setlocal
cd /d "%~dp0"
rem Kiểm tra độ sắc nét và co giãn của Bảng điều khiển trên từng màn hình. Cửa sổ thử hiện vài giây rồi tự đóng.
set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
if not exist "%FW%\csc.exe" set FW=%WINDIR%\Microsoft.NET\Framework\v4.0.30319
if not exist bin mkdir bin
"%FW%\csc.exe" /nologo /codepage:65001 /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /r:System.Security.dll /r:System.Xml.dll /out:bin\pane-shot.exe src\*.cs tests\MockWord.cs tools\PaneShot.cs
if errorlevel 1 exit /b 1
echo --- Man hinh 1 (chinh):
bin\pane-shot.exe pane-man1.png pmv2 0 0 -1 pane
echo --- Man hinh 2 (neu co):
bin\pane-shot.exe pane-man2.png pmv2 0 1 -1 pane
echo --- Keo tu man hinh 1 sang man hinh 2:
bin\pane-shot.exe pane-keo.png pmv2 0 0 1 pane
echo --- Hop Cai dat o man hinh 2:
bin\pane-shot.exe cai-dat-man2.png pmv2 0 1 -1 settings
echo Xem cac tep png vua tao. Dong "dpi=... nhanBietDPI=2" la dung (2 = sac net theo tung man hinh).
endlocal