@echo off
setlocal
cd /d "%~dp0"
set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
if not exist "%FW%\csc.exe" set FW=%WINDIR%\Microsoft.NET\Framework\v4.0.30319
if not exist bin mkdir bin
set COMMON=src\NumberUtil.cs src\Settings.cs src\Log.cs src\WordPalette.cs
"%FW%\csc.exe" /nologo /codepage:65001 /r:Microsoft.CSharp.dll /r:System.Core.dll /out:bin\tests-core.exe %COMMON% tests\TestMain.cs
if errorlevel 1 (pause & exit /b 1)
"%FW%\csc.exe" /nologo /codepage:65001 /r:Microsoft.CSharp.dll /r:System.Core.dll /out:bin\tests-num.exe %COMMON% tests\TestNum.cs
if errorlevel 1 (pause & exit /b 1)
"%FW%\csc.exe" /nologo /codepage:65001 /r:Microsoft.CSharp.dll /r:System.Core.dll /r:System.Xml.dll /out:bin\tests-review.exe %COMMON% src\Chem.cs src\Units.cs src\Review.cs src\Omml.cs src\MathSpans.cs src\Compose.cs src\NumberEdit.cs src\NumberWords.cs src\UnitWords.cs tests\TestReview.cs
if errorlevel 1 (pause & exit /b 1)
"%FW%\csc.exe" /nologo /codepage:65001 /r:Microsoft.CSharp.dll /r:System.Core.dll /r:System.Xml.dll /out:bin\tests-word.exe %COMMON% src\Chem.cs src\Units.cs src\Review.cs src\Omml.cs src\MathSpans.cs src\Compose.cs src\NumberEdit.cs src\NumberWords.cs src\UnitWords.cs src\WordTables.cs src\WordLayout.cs src\WordReview.cs src\AiReview.cs src\TranslationNumbers.cs src\WordFormat.cs src\AiClient.cs src\Secret.cs src\WordText.cs tests\MockWord.cs tests\TestWord.cs
if errorlevel 1 (pause & exit /b 1)
"%FW%\csc.exe" /nologo /codepage:65001 /r:Microsoft.CSharp.dll /r:System.Core.dll /r:System.Xml.dll /out:bin\tests-reviewword.exe %COMMON% src\Chem.cs src\Units.cs src\Review.cs src\Omml.cs src\MathSpans.cs src\Compose.cs src\NumberEdit.cs src\NumberWords.cs src\UnitWords.cs src\WordTables.cs src\WordReview.cs src\AiReview.cs src\TranslationNumbers.cs src\WordText.cs src\AiClient.cs src\Secret.cs src\DpiUtil.cs src\WordFormat.cs tests\MockWord.cs tests\TestReviewWord.cs
if errorlevel 1 (pause & exit /b 1)
"%FW%\csc.exe" /nologo /codepage:65001 /r:Microsoft.CSharp.dll /r:System.Core.dll /r:System.Xml.dll /out:bin\tests-numwords.exe %COMMON% src\Chem.cs src\Units.cs src\Review.cs src\Omml.cs src\MathSpans.cs src\Compose.cs src\NumberEdit.cs src\NumberWords.cs src\UnitWords.cs tests\TestNumWords.cs
if errorlevel 1 (pause & exit /b 1)
"%FW%\csc.exe" /nologo /codepage:65001 /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Security.dll /r:System.Xml.dll /r:Microsoft.CSharp.dll /out:bin\tests-ai.exe %COMMON% src\Secret.cs src\DpiUtil.cs src\Chem.cs src\Units.cs src\Review.cs src\Omml.cs src\MathSpans.cs src\Compose.cs src\WordTables.cs src\AiClient.cs src\AiReview.cs src\TranslationNumbers.cs src\WordText.cs tests\MockWord.cs tests\TestAi.cs
if errorlevel 1 (pause & exit /b 1)
"%FW%\csc.exe" /nologo /codepage:65001 /target:exe /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Core.dll /r:Microsoft.CSharp.dll /r:System.Xml.dll /out:bin\tests-colorpick.exe src\ColorPickerForm.cs src\DpiUtil.cs src\WordPalette.cs src\Log.cs src\Settings.cs src\NumberUtil.cs tests\TestColorPick.cs
if errorlevel 1 (pause & exit /b 1)
"%FW%\csc.exe" /nologo /codepage:65001 /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /r:System.Security.dll /r:System.Xml.dll /out:bin\tests-ribbon.exe src\*.cs tests\TestRibbon.cs
if errorlevel 1 (pause & exit /b 1)
"%FW%\csc.exe" /nologo /codepage:65001 /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /r:System.Security.dll /r:System.Xml.dll /out:bin\tests-icons.exe src\*.cs tests\TestIcons.cs
if errorlevel 1 (pause & exit /b 1)
"%FW%\csc.exe" /nologo /codepage:65001 /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /r:System.Security.dll /r:System.Xml.dll /out:bin\tests-dpi.exe src\*.cs tests\MockWord.cs tests\TestDpi.cs
if errorlevel 1 (pause & exit /b 1)
"%FW%\csc.exe" /nologo /codepage:65001 /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /r:System.Security.dll /r:System.Xml.dll /main:TestComposePanel /out:bin\tests-composepanel.exe src\*.cs tests\MockWord.cs tests\TestReviewWord.cs tests\TestComposePanel.cs
if errorlevel 1 exit /b 1
chcp 65001 >nul
call :verify tests-core
if errorlevel 1 exit /b 1
call :verify tests-num
if errorlevel 1 exit /b 1
call :verify tests-review
if errorlevel 1 exit /b 1
call :verify tests-word
if errorlevel 1 exit /b 1
call :verify tests-reviewword
if errorlevel 1 exit /b 1
call :verify tests-numwords
if errorlevel 1 exit /b 1
call :verify tests-ai
if errorlevel 1 exit /b 1
call :verify tests-colorpick
if errorlevel 1 exit /b 1
call :verify tests-ribbon
if errorlevel 1 exit /b 1
call :verify tests-icons
if errorlevel 1 exit /b 1
call :verify tests-dpi
if errorlevel 1 exit /b 1
echo.
call :verify tests-composepanel
if errorlevel 1 exit /b 1
echo Da chay xong cac bo kiem thu.
pause
endlocal
exit /b 0

:verify
"bin\%~1.exe" >"bin\%~1.log" 2>&1
set RESULT=%errorlevel%
findstr /v /b "PASS" "bin\%~1.log"
exit /b %RESULT%
