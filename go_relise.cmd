@echo off
echo Building release version...

if not exist "Release" mkdir Release

dotnet publish -c Release ^
    -r win-x64 ^
    --self-contained true ^
    -o "Release" ^
    -p:PublishSingleFile=true ^
    -p:PublishTrimmed=false ^
    -p:IncludeNativeLibrariesForSelfExtract=true ^
    -p:IncludeAllContentForSelfExtract=true ^
    -p:DebugType=None ^
    -p:DebugSymbols=false ^
    -p:Optimize=true ^
    -p:EnableCompressionInSingleFile=true ^
    -p:SelfContained=true ^
    -p:PublishReadyToRun=true

echo Build completed!
if exist "Учёт заявок.xlsx" copy "Учёт заявок.xlsx" Release\
echo Preparation completed!
pause