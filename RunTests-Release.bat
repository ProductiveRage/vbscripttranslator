@echo off

%~d0
cd "%~p0"
packages\xunit.runner.console.2.1.0\tools\xunit.console.x86.exe UnitTests\bin\Release\VBScriptTranslator.UnitTests.dll -noappdomain

echo.
pause