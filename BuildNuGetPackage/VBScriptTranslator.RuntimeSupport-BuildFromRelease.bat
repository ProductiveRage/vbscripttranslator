@echo off

%~d0
cd "%~p0"

del *.nu*
del *.dll
del *.pdb
del *.xml

copy ..\RuntimeSupport\bin\Release\* > nul
copy ..\VBScriptTranslator.RuntimeSupport.nuspec > nul

..\packages\NuGet.CommandLine.3.3.0\tools\nuget pack -NoPackageAnalysis VBScriptTranslator.RuntimeSupport.nuspec