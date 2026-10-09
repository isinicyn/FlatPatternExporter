@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

:: FlatPatternExporter - GitHub release creator
::
:: Usage: create-release-draft.bat [notes.md] [--publish] [--yes]
::   notes.md   Release notes in Markdown; {VERSION} is replaced with the full version.
::              Without it a template is used.
::   --publish  Publish the release immediately instead of creating a draft.
::   --yes      Do not ask questions; stop with an error instead of replacing an existing tag or release.
::
:: Archives for the current version are taken from Release\ (build them with "publish.bat all").

set "EXIT_CODE=0"
set "NOTES_SOURCE="
set "PUBLISH=0"
set "ASSUME_YES=0"

:ParseArgs
if "%~1"=="" goto :ArgsDone
if /i "%~1"=="--publish" (
    set "PUBLISH=1"
) else if /i "%~1"=="--yes" (
    set "ASSUME_YES=1"
) else if /i "%~1"=="--help" (
    goto :Usage
) else if defined NOTES_SOURCE (
    echo [ERROR] Unexpected argument: %~1
    goto :Usage
) else (
    set "NOTES_SOURCE=%~f1"
)
shift
goto :ParseArgs
:ArgsDone

echo ========================================
echo FlatPatternExporter - Release Creator
echo ========================================
echo.

where gh >nul 2>&1
if errorlevel 1 (
    echo [ERROR] GitHub CLI ^(gh^) not found. Install it from https://cli.github.com/
    goto :Fail
)

:: Version: VersionPrefix from .csproj + Git commit count
set "VERSION_PREFIX="
for /f "tokens=3 delims=<>" %%i in ('findstr /C:"<VersionPrefix>" FlatPatternExporter\FlatPatternExporter.csproj') do set "VERSION_PREFIX=%%i"
if not defined VERSION_PREFIX (
    echo [ERROR] Could not read VersionPrefix from FlatPatternExporter\FlatPatternExporter.csproj
    goto :Fail
)
for /f %%i in ('git rev-list --count HEAD') do set "COMMIT_COUNT=%%i"
set "VERSION=%VERSION_PREFIX%.%COMMIT_COUNT%"
set "TAG=v%VERSION%"
set "TITLE=Flat Pattern Exporter %VERSION_PREFIX%"

:: The tag must point to a commit that is already on the remote branch
git fetch --quiet origin
set "UNPUSHED="
for /f %%i in ('git rev-list --count @{u}..HEAD 2^>nul') do set "UNPUSHED=%%i"
if not defined UNPUSHED (
    echo [ERROR] The current branch has no upstream branch.
    goto :Fail
)
if not "%UNPUSHED%"=="0" (
    echo [ERROR] %UNPUSHED% commit^(s^) are not pushed yet. Push them before creating the release.
    goto :Fail
)

:: Archives of exactly this version
set "ARCHIVES="
set "MISSING=0"
for %%a in (
    "Release\FlatPatternExporter-%TAG%-x64-Deploy.zip"
    "Release\FlatPatternExporter-%TAG%-x64-Portable.zip"
    "Release\FlatPatternExporter-%TAG%-x64-FrameworkDependent.zip"
    "Release\FlatPatternExporter.Updater-%TAG%-x64.zip"
) do (
    if exist %%a (
        set "ARCHIVES=!ARCHIVES! %%a"
    ) else (
        echo [ERROR] Missing archive: %%~a
        set "MISSING=1"
    )
)
if "%MISSING%"=="1" (
    echo Build the archives for %TAG% with "publish.bat all".
    goto :Fail
)

:: Release notes
set "NOTES_FILE=%TEMP%\fpe_release_notes_%VERSION%.md"
if defined NOTES_SOURCE (
    if not exist "%NOTES_SOURCE%" (
        echo [ERROR] Release notes file not found: %NOTES_SOURCE%
        goto :Fail
    )
    powershell -NoProfile -Command "$t = [IO.File]::ReadAllText('%NOTES_SOURCE%').Replace('{VERSION}', '%VERSION%'); [IO.File]::WriteAllText('%NOTES_FILE%', $t, (New-Object Text.UTF8Encoding $false))"
    if errorlevel 1 (
        echo [ERROR] Could not prepare release notes
        goto :Fail
    )
) else (
    > "%NOTES_FILE%" (
        echo Describe the release in one sentence.
        echo.
        echo ## What's new
        echo.
        echo -
        echo.
        echo ## Improvements
        echo.
        echo -
        echo.
        echo ## Fixes
        echo.
        echo -
        echo.
        echo ## Downloads
        echo.
        echo - Portable: single .exe, no installation.
        echo - Deploy: for installers ^(Inno Setup, WiX, NSIS^).
        echo - FrameworkDependent: requires .NET 8 Desktop Runtime, smallest download.
        echo - Updater: used by the built-in updater, no need to download it.
    )
)

if "%PUBLISH%"=="1" (set "MODE=published") else (set "MODE=draft")
echo Release:  %TAG%
echo Title:    %TITLE%
echo Mode:     %MODE%
if defined NOTES_SOURCE (echo Notes:    %NOTES_SOURCE%) else (echo Notes:    template)
echo Archives:
for %%a in (%ARCHIVES%) do echo   - %%~nxa
echo.

if "%ASSUME_YES%"=="0" (
    set "CONFIRM="
    set /p "CONFIRM=Create %MODE% release %TAG%? (y/n): "
    if /i not "!CONFIRM!"=="y" (
        echo Cancelled.
        goto :Cleanup
    )
)

:: Existing tag or release
set "TAG_LOCAL=0"
set "TAG_REMOTE=0"
set "RELEASE_EXISTS=0"
git rev-parse -q --verify "refs/tags/%TAG%" >nul 2>&1 && set "TAG_LOCAL=1"
git ls-remote --exit-code --tags origin "refs/tags/%TAG%" >nul 2>&1 && set "TAG_REMOTE=1"
gh release view %TAG% >nul 2>&1 && set "RELEASE_EXISTS=1"

if "%TAG_LOCAL%%TAG_REMOTE%%RELEASE_EXISTS%" neq "000" (
    echo [WARNING] %TAG% already exists: local tag=%TAG_LOCAL%, remote tag=%TAG_REMOTE%, release=%RELEASE_EXISTS%
    if "%ASSUME_YES%"=="1" (
        echo [ERROR] Delete the existing tag and release or run the script without --yes.
        goto :Fail
    )
    set "REPLACE="
    set /p "REPLACE=Delete them and create the release again? (y/n): "
    if /i not "!REPLACE!"=="y" (
        echo Cancelled.
        goto :Cleanup
    )
    if "%RELEASE_EXISTS%"=="1" (
        gh release delete %TAG% --yes || goto :Fail
    )
    if "%TAG_REMOTE%"=="1" (
        git push origin --delete %TAG% || goto :Fail
    )
    if "%TAG_LOCAL%"=="1" (
        git tag -d %TAG% || goto :Fail
    )
)

echo [1/3] Creating tag %TAG%...
git tag -a %TAG% -m "Release %VERSION%"
if errorlevel 1 goto :Fail

echo [2/3] Pushing tag...
git push origin %TAG%
if errorlevel 1 goto :Fail

echo [3/3] Creating %MODE% release and uploading archives...
if "%PUBLISH%"=="1" (set "RELEASE_FLAGS=--latest") else (set "RELEASE_FLAGS=--draft")
gh release create %TAG% %RELEASE_FLAGS% --verify-tag --title "%TITLE%" --notes-file "%NOTES_FILE%" %ARCHIVES%
if errorlevel 1 (
    echo [ERROR] Failed to create the release. The tag %TAG% is already pushed;
    echo         rerun the script to replace it, or remove it: git push origin --delete %TAG% ^&^& git tag -d %TAG%
    goto :Fail
)

set "RELEASE_URL="
for /f "tokens=*" %%i in ('gh release view %TAG% --json url -q .url') do set "RELEASE_URL=%%i"
echo.
echo ========================================
echo Release %TAG% created ^(%MODE%^)
echo ========================================
echo %RELEASE_URL%
if "%PUBLISH%"=="0" echo Review the notes and publish it: %RELEASE_URL:/tag/=/edit/%
goto :Cleanup

:Usage
echo Usage: create-release-draft.bat [notes.md] [--publish] [--yes]
echo   notes.md   Release notes in Markdown; {VERSION} is replaced with the full version
echo   --publish  Publish the release immediately instead of creating a draft
echo   --yes      Do not ask questions; stop with an error instead of replacing an existing tag or release
set "EXIT_CODE=1"
goto :End

:Fail
set "EXIT_CODE=1"

:Cleanup
if defined NOTES_FILE if exist "%NOTES_FILE%" del "%NOTES_FILE%" >nul 2>&1

:End
echo.
if "%ASSUME_YES%"=="0" pause
endlocal & exit /b %EXIT_CODE%
