@echo off
setlocal enabledelayedexpansion

echo ========================================
echo FlatPatternExporter - Release Draft Creator
echo ========================================
echo.

:: Check if gh CLI is available
where gh >nul 2>&1
if errorlevel 1 (
    echo [ERROR] GitHub CLI (gh^) not found. Please install it from https://cli.github.com/
    pause
    exit /b 1
)

:: Extract VersionPrefix from .csproj
echo [INFO] Extracting version from project file...
for /f "tokens=3 delims=<>" %%i in ('findstr /C:"<VersionPrefix>" FlatPatternExporter\FlatPatternExporter.csproj') do set VERSION_PREFIX=%%i

if not defined VERSION_PREFIX (
    echo [ERROR] Could not extract VersionPrefix from .csproj file
    pause
    exit /b 1
)

:: Get revision number from Git commits
for /f "tokens=*" %%i in ('git rev-list --count HEAD') do set COMMIT_COUNT=%%i

:: Combine into full version
set VERSION=%VERSION_PREFIX%.%COMMIT_COUNT%

echo [INFO] Detected version: v%VERSION%
echo [INFO] Version prefix: %VERSION_PREFIX%
echo [INFO] Git commit count: %COMMIT_COUNT%
echo.

:: Check if Release directory exists and has archives
if exist "Release\" (
    set RELEASE_EXISTS=1
) else (
    echo [ERROR] Release directory not found. Run publish.bat first.
    pause
    exit /b 1
)

:: Count zip files
set ZIP_COUNT=0
for %%f in (Release\*.zip) do set /a ZIP_COUNT+=1

if %ZIP_COUNT% equ 0 (
    echo [ERROR] No .zip archives found in Release\ directory.
    echo Please run publish.bat to create release archives.
    pause
    exit /b 1
)

echo [INFO] Found %ZIP_COUNT% archive(s^) in Release\ directory
echo.

:: List archives
echo Archives to upload:
for %%f in (Release\*.zip) do (
    echo   - %%~nxf
)
echo.

:: Verify archive version matches calculated version
echo [INFO] Verifying archive versions...
set ARCHIVE_VERSION_MISMATCH=0
for %%f in (Release\FlatPatternExporter-v*.zip) do (
    set ARCHIVE_NAME=%%~nf
    for /f "tokens=2 delims=-" %%v in ("!ARCHIVE_NAME!") do (
        set ARCHIVE_VERSION=%%v
        if not "!ARCHIVE_VERSION!"=="v%VERSION%" (
            echo [WARNING] Archive version mismatch: %%~nxf has !ARCHIVE_VERSION!, expected v%VERSION%
            set ARCHIVE_VERSION_MISMATCH=1
        )
    )
    goto :version_check_done
)
:version_check_done

if %ARCHIVE_VERSION_MISMATCH% equ 1 (
    echo.
    echo [ERROR] Archive version mismatch detected!
    echo Expected version: v%VERSION%
    echo.
    echo This usually means archives were built with a different commit count.
    echo Please run publish.bat again to rebuild archives with the current version.
    pause
    exit /b 1
)
echo [SUCCESS] Archive versions match: v%VERSION%
echo.

:: Confirm action
set /p CONFIRM="Create draft release v%VERSION%? (y/n): "
if /i "%CONFIRM%" neq "y" (
    echo Cancelled.
    exit /b 0
)
echo.

:: Check if tag already exists (local and remote)
set TAG_EXISTS_LOCAL=0
set TAG_EXISTS_REMOTE=0

git rev-parse v%VERSION% >nul 2>&1
if %errorlevel% equ 0 set TAG_EXISTS_LOCAL=1

git ls-remote --tags origin v%VERSION% >nul 2>&1
if %errorlevel% equ 0 set TAG_EXISTS_REMOTE=1

if %TAG_EXISTS_LOCAL% equ 1 (
    echo [WARNING] Tag v%VERSION% already exists locally
)
if %TAG_EXISTS_REMOTE% equ 1 (
    echo [WARNING] Tag v%VERSION% already exists on remote
)

if %TAG_EXISTS_LOCAL% equ 1 (
    set /p DELETE_TAG="Delete and recreate tag? (y/n): "
    if /i "!DELETE_TAG!" neq "y" (
        echo Cancelled.
        exit /b 0
    )

    if %TAG_EXISTS_LOCAL% equ 1 (
        echo [INFO] Deleting local tag v%VERSION%...
        git tag -d v%VERSION%
        if errorlevel 1 (
            echo [ERROR] Failed to delete local tag
            pause
            exit /b 1
        )
    )

    if %TAG_EXISTS_REMOTE% equ 1 (
        echo [INFO] Deleting remote tag v%VERSION%...
        git push origin --delete v%VERSION%
        if errorlevel 1 (
            echo [ERROR] Failed to delete remote tag
            pause
            exit /b 1
        )
    )
) else (
    if %TAG_EXISTS_REMOTE% equ 1 (
        set /p DELETE_TAG="Delete remote tag and create new? (y/n): "
        if /i "!DELETE_TAG!" neq "y" (
            echo Cancelled.
            exit /b 0
        )

        echo [INFO] Deleting remote tag v%VERSION%...
        git push origin --delete v%VERSION%
        if errorlevel 1 (
            echo [ERROR] Failed to delete remote tag
            pause
            exit /b 1
        )
    )
)
echo.

:: Check if release already exists on GitHub
echo [INFO] Checking for existing release on GitHub...
gh release view v%VERSION% >nul 2>&1
if %errorlevel% equ 0 (
    echo [WARNING] Release v%VERSION% already exists on GitHub
    set /p DELETE_RELEASE="Delete existing release and create new? (y/n): "
    if /i "!DELETE_RELEASE!" neq "y" (
        echo Cancelled.
        exit /b 0
    )

    echo [INFO] Deleting existing release v%VERSION%...
    gh release delete v%VERSION% --yes
    if errorlevel 1 (
        echo [ERROR] Failed to delete existing release
        pause
        exit /b 1
    )
    echo [SUCCESS] Existing release deleted
) else (
    echo [INFO] No existing release found
)
echo.

:: Create tag
echo [STEP 1/4] Creating tag v%VERSION%...
git tag -a v%VERSION% -m "Release %VERSION%"
if errorlevel 1 (
    echo [ERROR] Failed to create tag
    pause
    exit /b 1
)
echo [SUCCESS] Tag created
echo.

:: Push tag
echo [STEP 2/4] Pushing tag to GitHub...
git push origin v%VERSION%
if errorlevel 1 (
    echo [ERROR] Failed to push tag
    pause
    exit /b 1
)
echo [SUCCESS] Tag pushed
echo.

:: Create draft release with archives
echo [STEP 3/4] Creating draft release on GitHub...
echo [INFO] This may take a few minutes to upload archives...
echo.

gh release create v%VERSION% --draft --title "v%VERSION%" --notes "## Release Notes%0A%0APlease add release notes here.%0A%0A### Archives%0A%0A- **Deploy** - For installers (Inno Setup, WiX, NSIS)%0A- **Portable** - Single executable, no installation required%0A- **FrameworkDependent** - Requires .NET 8.0 Runtime%0A- **Updater** - For automatic updates" Release\*.zip

if errorlevel 1 (
    echo [ERROR] Failed to create draft release
    echo You may need to delete the tag manually: git tag -d v%VERSION% ^&^& git push origin --delete v%VERSION%
    pause
    exit /b 1
)

echo [SUCCESS] Draft release created
echo.

:: Get release URL
echo [STEP 4/4] Getting release URL...
for /f "tokens=*" %%i in ('gh release view v%VERSION% --json url -q .url') do set RELEASE_URL=%%i

echo.
echo ========================================
echo Draft Release Created Successfully!
echo ========================================
echo.
echo Version: v%VERSION%
echo Tag: v%VERSION%
echo Status: DRAFT
echo.
echo Release URL:
echo %RELEASE_URL%
echo.
echo Edit URL:
echo %RELEASE_URL%/edit
echo.
echo [NEXT STEPS]
echo 1. Open the edit URL in your browser
echo 2. Update release notes with detailed description
echo 3. Click "Publish release" when ready
echo.
pause
