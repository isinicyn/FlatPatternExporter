@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

:: FlatPatternExporter - Publish Script
::
:: Usage: publish.bat [deploy|portable|framework|updater|all]
::   Without arguments an interactive menu is shown.
::   "all" removes old archives from Release\ and builds the full set for a GitHub release.
:: Archives are saved to Release\. On a failed build the dotnet output is kept in Release\publish-*.log.

set "EXIT_CODE=0"
set "INTERACTIVE=0"
set "BUILD_VERSION="

echo =====================================
echo FlatPatternExporter - Publish Script
echo =====================================
echo.

if "%~1"=="" (
    set "INTERACTIVE=1"
    call :ShowMenu
) else (
    call :Run %~1
)
goto :End

:ShowMenu
echo Select publish profile:
echo.
echo 1. Deploy     - Archive with separate DLLs for installers
echo 2. Portable   - Archive with a single .exe file
echo 3. Framework  - Archive that requires .NET 8 Runtime, minimal size
echo 4. Updater    - Updater archive only
echo 5. All        - Full set of archives for a GitHub release
echo 6. Exit
echo.
set "CHOICE="
set "TARGET="
set /p "CHOICE=Your choice (1-6): "
if "%CHOICE%"=="1" set "TARGET=deploy"
if "%CHOICE%"=="2" set "TARGET=portable"
if "%CHOICE%"=="3" set "TARGET=framework"
if "%CHOICE%"=="4" set "TARGET=updater"
if "%CHOICE%"=="5" set "TARGET=all"
if "%CHOICE%"=="6" goto :eof
if not defined TARGET (
    echo [ERROR] Invalid choice: %CHOICE%
    set "EXIT_CODE=1"
    goto :eof
)
call :Run %TARGET%
goto :eof

:Run
if /i "%~1"=="deploy" (
    call :PublishArchive DeployProfile deploy Deploy
) else if /i "%~1"=="portable" (
    call :PublishArchive PortableProfile portable Portable
) else if /i "%~1"=="framework" (
    call :PublishArchive FrameworkDependentProfile framework-dependent FrameworkDependent
) else if /i "%~1"=="updater" (
    call :PublishUpdater
) else if /i "%~1"=="all" (
    call :PublishAll
) else (
    echo [ERROR] Unknown profile: %~1
    echo Usage: publish.bat [deploy^|portable^|framework^|updater^|all]
    set "EXIT_CODE=1"
)
goto :eof

:PublishAll
echo [ALL] Publishing the full set of archives...
if exist "Release\*.zip" (
    echo [CLEAN] Removing old archives from Release\
    del /Q "Release\*.zip"
)
call :PublishArchive DeployProfile deploy Deploy
if "!EXIT_CODE!"=="1" goto :eof
call :PublishArchive PortableProfile portable Portable
if "!EXIT_CODE!"=="1" goto :eof
call :PublishArchive FrameworkDependentProfile framework-dependent FrameworkDependent
if "!EXIT_CODE!"=="1" goto :eof
call :PublishUpdater
if "!EXIT_CODE!"=="1" goto :eof
echo.
echo [SUCCESS] All archives for v%BUILD_VERSION% are ready in Release\
goto :eof

:: %1 - publish profile, %2 - publish folder, %3 - build type (archive suffix and .buildtype marker)
:PublishArchive
echo.
echo [%~3] Publishing...
call :PublishMainApp %~1 %~2
if "!EXIT_CODE!"=="1" goto :eof

call :PrepareStaging
if /i "%~2"=="portable" (
    copy /Y "FlatPatternExporter\bin\publish\portable\FlatPatternExporter.exe" "Release\staging\" >nul
) else (
    xcopy "FlatPatternExporter\bin\publish\%~2\*" "Release\staging\" /E /I /Y /Q >nul
)
if errorlevel 1 (
    echo [ERROR] Failed to copy published files
    set "EXIT_CODE=1"
    goto :eof
)
echo %~3> "Release\staging\.buildtype"
call :Zip "Release\staging\*" "Release\FlatPatternExporter-v%BUILD_VERSION%-x64-%~3.zip"
goto :eof

:PublishUpdater
echo.
echo [Updater] Publishing...
:: The archive name uses the application version, so the main application has to be built first
if not defined BUILD_VERSION (
    call :PublishMainApp PortableProfile portable
    if "!EXIT_CODE!"=="1" goto :eof
)
call :DotnetPublish "FlatPatternExporter.Updater\FlatPatternExporter.Updater.csproj" PortableProfile "FlatPatternExporter.Updater\bin\publish\portable" "Release\publish-updater.log"
if "!EXIT_CODE!"=="1" goto :eof

call :PrepareStaging
copy /Y "FlatPatternExporter.Updater\bin\publish\portable\FlatPatternExporter.Updater.exe" "Release\staging\" >nul
if errorlevel 1 (
    echo [ERROR] Failed to copy the updater
    set "EXIT_CODE=1"
    goto :eof
)
call :Zip "Release\staging\*.exe" "Release\FlatPatternExporter.Updater-v%BUILD_VERSION%-x64.zip"
goto :eof

:: %1 - publish profile, %2 - publish folder; sets BUILD_VERSION
:PublishMainApp
set "PUBLISH_LOG=Release\publish-%~1.log"
call :DotnetPublish "FlatPatternExporter\FlatPatternExporter.csproj" %~1 "FlatPatternExporter\bin\publish\%~2" "%PUBLISH_LOG%" keep
if "!EXIT_CODE!"=="1" goto :eof

set "BUILD_VERSION="
for /f "tokens=3" %%v in ('findstr /C:"File Version:" "%PUBLISH_LOG%"') do set "BUILD_VERSION=%%v"
del "%PUBLISH_LOG%" >nul 2>&1
if not defined BUILD_VERSION (
    echo [ERROR] Could not detect the application version from the build output
    set "EXIT_CODE=1"
    goto :eof
)
echo [VERSION] %BUILD_VERSION%
goto :eof

:: %1 - project, %2 - publish profile, %3 - publish folder (cleaned first), %4 - log file, %5 - "keep" to leave the log for the caller
:DotnetPublish
if not exist "Release" mkdir "Release"
if exist "%~3" rmdir /S /Q "%~3"
echo Running dotnet publish ^(%~2^)...
dotnet publish "%~1" --configuration Release /p:PublishProfile=%~2 > "%~4" 2>&1
if errorlevel 1 (
    echo [ERROR] dotnet publish failed. Last lines of %~4:
    powershell -NoProfile -Command "Get-Content -Tail 25 '%~4'"
    set "EXIT_CODE=1"
    goto :eof
)
if /i not "%~5"=="keep" del "%~4" >nul 2>&1
goto :eof

:PrepareStaging
if exist "Release\staging" rmdir /S /Q "Release\staging"
mkdir "Release\staging"
goto :eof

:: %1 - files to pack, %2 - archive path
:Zip
if exist "%~2" del "%~2"
powershell -NoProfile -Command "Compress-Archive -Path '%~1' -DestinationPath '%~2' -CompressionLevel Optimal"
if errorlevel 1 (
    echo [ERROR] Failed to create %~nx2
    set "EXIT_CODE=1"
    goto :eof
)
rmdir /S /Q "Release\staging" >nul 2>&1
echo [SUCCESS] Release\%~nx2
goto :eof

:End
echo.
if "%EXIT_CODE%"=="0" (echo Script completed.) else (echo Script failed.)
if "%INTERACTIVE%"=="1" pause
endlocal & exit /b %EXIT_CODE%
