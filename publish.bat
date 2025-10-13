@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

:: FlatPatternExporter Publish Script
:: This script automates project publishing for various deployment scenarios

echo =====================================
echo FlatPatternExporter - Publish Script
echo =====================================
echo.

:: Check command line parameters
if "%1"=="" (
    call :ShowMenu
) else (
    call :PublishProfile %1
)
goto :End

:ShowMenu
echo Select publish profile:
echo.
echo 1. Deploy          - Ready files for installer (separate DLLs)
echo 2. Portable        - Portable version (single .exe file)
echo 3. Framework       - Depends on .NET 8 Runtime (minimal size)
echo 4. GitHub Release  - Create zip archive for GitHub Release
echo 5. All             - Publish all profiles
echo 6. Exit            - Exit
echo.
set /p choice="Your choice (1-6): "

if "%choice%"=="1" call :PublishProfile deploy
if "%choice%"=="2" call :PublishProfile portable
if "%choice%"=="3" call :PublishProfile framework
if "%choice%"=="4" call :PublishGitHubRelease
if "%choice%"=="5" call :PublishAllProfiles
if "%choice%"=="6" exit /b 0

goto :eof

:PublishProfile
set profile=%1

if "%profile%"=="deploy" (
    echo.
    echo [DEPLOY] Publishing for installer...
    echo.
    call :PublishMainApp DeployProfile deploy
    call :CopyToReleaseFolder deploy Deploy
)

if "%profile%"=="portable" (
    echo.
    echo [PORTABLE] Publishing portable version...
    echo.
    call :PublishMainApp PortableProfile portable
    call :CopyToReleaseFolder portable Portable
)

if "%profile%"=="framework" (
    echo.
    echo [FRAMEWORK] Publishing framework-dependent version...
    echo.
    call :PublishMainApp FrameworkDependentProfile framework-dependent
    call :CopyToReleaseFolder framework-dependent FrameworkDependent
)

goto :eof

:PublishAllProfiles
echo.
echo [ALL] Publishing all profiles...
echo.
call :PublishProfile deploy
echo.
call :PublishProfile portable
echo.
call :PublishProfile framework
echo.
call :PublishGitHubRelease
echo.
echo [SUCCESS] All profiles published!
goto :eof

:PublishGitHubRelease
echo.
echo [GITHUB RELEASE] Creating GitHub Release archive...
echo.

echo [1/3] Publishing main application (Portable)...
call :PublishMainApp PortableProfile portable

if errorlevel 1 (
    echo [ERROR] Main application publishing failed
    goto :eof
)

echo [2/3] Publishing Updater (Portable)...
dotnet publish "FlatPatternExporter.Updater\FlatPatternExporter.Updater.csproj" ^
    --configuration Release ^
    /p:PublishProfile=PortableProfile >nul 2>&1

if errorlevel 1 (
    echo [ERROR] Updater publishing failed
    goto :eof
)
echo [SUCCESS] Updater published

echo [3/3] Creating release archive...

if not defined BUILD_VERSION (
    echo [ERROR] Could not create release archive - version not defined
    goto :eof
)

:: Create Release\GitHub directory
set githubDir=Release\GitHub
if not exist "%githubDir%" mkdir "%githubDir%"

:: Copy exe files to GitHub directory
copy "FlatPatternExporter\bin\publish\portable\FlatPatternExporter.exe" "%githubDir%\" /Y >nul 2>&1
if not errorlevel 1 echo [SUCCESS] Main application copied

copy "FlatPatternExporter.Updater\bin\publish\portable\FlatPatternExporter.Updater.exe" "%githubDir%\" /Y >nul 2>&1
if not errorlevel 1 echo [SUCCESS] Updater copied

:: Create zip archive
set archiveName=FlatPatternExporter-v%BUILD_VERSION%.zip
set archivePath=%githubDir%\%archiveName%

if exist "%archivePath%" del "%archivePath%" >nul 2>&1

powershell -NoProfile -Command "Compress-Archive -Path '%githubDir%\*.exe' -DestinationPath '%archivePath%' -CompressionLevel Optimal" >nul 2>&1

if not errorlevel 1 (
    echo [SUCCESS] GitHub Release archive created: %archiveName%

    :: Delete source exe files, keep only zip
    del "%githubDir%\*.exe" >nul 2>&1

    echo.
    echo [INFO] Archive ready for GitHub Release: %githubDir%\%archiveName%
    echo [INFO] Files in archive: FlatPatternExporter.exe, FlatPatternExporter.Updater.exe
) else (
    echo [ERROR] Failed to create release archive
)

echo.
goto :eof

:PublishMainApp
set profileName=%1
set outputFolder=%2

echo Publishing main application...

:: Publish and extract version from output stream (suppress output)
for /f "delims=" %%i in ('dotnet publish "FlatPatternExporter\FlatPatternExporter.csproj" --configuration Release /p:PublishProfile^=%profileName% 2^>^&1 ^| findstr /C:"File Version:"') do (
    for /f "tokens=3" %%v in ("%%i") do set BUILD_VERSION=%%v
)

if defined BUILD_VERSION (
    echo [VERSION] Detected version: %BUILD_VERSION%
    echo [SUCCESS] Main application published
) else (
    echo [ERROR] Main application publishing failed - version not detected
    exit /b 1
)

goto :eof

:CopyToReleaseFolder
set sourceFolder=%1
set targetFolder=%2

echo.
echo [COPY] Copying files to Release\%targetFolder%...

:: Create target folder
if not exist "Release\%targetFolder%" mkdir "Release\%targetFolder%"

:: Copy files based on profile
if "%sourceFolder%"=="portable" (
    :: For Portable profile copy only .exe files ^(SingleFile^)
    copy "FlatPatternExporter\bin\publish\%sourceFolder%\FlatPatternExporter.exe" "Release\%targetFolder%\" /Y >nul 2>&1
    if not errorlevel 1 echo [SUCCESS] Main application copied ^(SingleFile^)
) else (
    :: For Deploy and FrameworkDependent copy all files
    xcopy "FlatPatternExporter\bin\publish\%sourceFolder%\*" "Release\%targetFolder%\" /E /I /Y >nul 2>&1
    if not errorlevel 1 echo [SUCCESS] Main application copied
)

:: Extract version from compiled exe and create version file
call :CreateVersionFile "%targetFolder%"

echo.
echo [INFO] Ready files are in: Release\%targetFolder%\
echo.
goto :eof

:CreateVersionFile
set targetFolder=%~1

if defined BUILD_VERSION (
    :: Create empty version file
    set versionFile=Release\%targetFolder%\%BUILD_VERSION%
    type nul > "!versionFile!"

    if not errorlevel 1 (
        echo [SUCCESS] Version file created: %BUILD_VERSION%
    ) else (
        echo [WARNING] Failed to create version file
    )
) else (
    echo [WARNING] Could not extract version from build output
)
goto :eof

:End
echo.
echo Script completed.
pause
endlocal
