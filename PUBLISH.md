# FlatPatternExporter Publishing Guide

This document describes the application publishing process for various deployment scenarios.

## Publish Profiles

### 1. **Deploy Profile** (Recommended for installers)
Publishes the application with separate DLL files.

**Main Application:**
- File: `FlatPatternExporter\Properties\PublishProfiles\DeployProfile.pubxml`
- Output: `FlatPatternExporter\bin\publish\deploy\`

**Characteristics:**
- ✅ Self-contained (.NET Runtime included)
- ✅ Separate DLL files
- ✅ PublishReadyToRun (startup optimization)
- ✅ Transparent file structure
- 📦 Size: ~100-120 MB
- 🎯 Use case: Inno Setup, WiX, NSIS installers

**Manual publishing:**
```bash
dotnet publish FlatPatternExporter\FlatPatternExporter.csproj --configuration Release /p:PublishProfile=DeployProfile
```

---

### 2. **Portable Profile** (For ZIP archives)
Publishes the application as a single executable file.

**Main Application:**
- File: `FlatPatternExporter\Properties\PublishProfiles\PortableProfile.pubxml`
- Output: `FlatPatternExporter\bin\publish\portable\`

**Characteristics:**
- ✅ Self-contained (.NET Runtime included)
- ✅ PublishSingleFile (single .exe)
- ✅ EnableCompressionInSingleFile
- ✅ IncludeNativeLibrariesForSelfExtract
- 📦 Size: ~150 MB
- 🎯 Use case: Portable version, ZIP archives

**Manual publishing:**
```bash
dotnet publish FlatPatternExporter\FlatPatternExporter.csproj --configuration Release /p:PublishProfile=PortableProfile
```

---

### 3. **FrameworkDependent Profile** (Minimal size)
Publishes the application without including .NET Runtime.

**Main Application:**
- File: `FlatPatternExporter\Properties\PublishProfiles\FrameworkDependentProfile.pubxml`
- Output: `FlatPatternExporter\bin\publish\framework-dependent\`

**Characteristics:**
- ⚠️ Requires .NET 8.0 Runtime on target machine
- ✅ Minimal size
- ✅ Separate DLL files
- 📦 Size: ~5-10 MB
- 🎯 Use case: Corporate environments with centralized .NET Runtime management

**Manual publishing:**
```bash
dotnet publish FlatPatternExporter\FlatPatternExporter.csproj --configuration Release /p:PublishProfile=FrameworkDependentProfile
```

---

### 4. **Updater Portable Profile** (For deployment)
Publishes only the updater application as a portable single-file executable.

**Updater Application:**
- File: `FlatPatternExporter.Updater\Properties\PublishProfiles\PortableProfile.pubxml`
- Output: `FlatPatternExporter.Updater\bin\publish\portable\`

**Characteristics:**
- ✅ Self-contained (.NET Runtime included)
- ✅ PublishSingleFile (single .exe)
- ✅ Creates zip archive with updater only
- ✅ Archive saved to `Release\`
- 📦 Archive size: ~80 MB
- 🎯 Use case: Deployment of updater for automatic updates

**Archive structure:**
```
FlatPatternExporter.Updater-v3.0.0-x64.zip
└── FlatPatternExporter.Updater.exe  # Updater
```

**Process:**
1. Publishes updater (Portable profile)
2. Creates zip archive with updater
3. Saves to `Release\FlatPatternExporter.Updater-v{VERSION}-x64.zip`
4. Automatic cleanup of temporary files

---

## Automated Publishing

### Using `publish.bat`

The `publish.bat` script automates the publishing process for all profiles.

#### Interactive mode:
```bash
publish.bat
```

Menu:
```
1. Deploy     - Archive with separate DLLs for installers
2. Portable   - Archive with a single .exe file
3. Framework  - Archive that requires .NET 8 Runtime, minimal size
4. Updater    - Updater archive only
5. All        - Full set of archives for a GitHub release
6. Exit
```

#### Command line mode:
```bash
publish.bat deploy      # Deploy archive
publish.bat portable    # Portable archive
publish.bat framework   # Framework-dependent archive
publish.bat updater     # Updater archive
publish.bat all         # Full set for a release; old archives in Release\ are removed first
```

In command line mode the script does not wait for a key press and returns exit code `1` on failure.
When `dotnet publish` fails, its last lines are printed and the full output stays in `Release\publish-*.log`.

### What does the script do?

**For Deploy, Portable, Framework profiles:**
1. **Publishes** main application with the selected profile
2. **Creates** `.buildtype` marker file (Deploy/Portable/FrameworkDependent)
3. **Creates** zip archive with all files
4. **Saves** to `Release\FlatPatternExporter-v{VERSION}-x64-{BUILD_TYPE}.zip`
5. **Cleans up** temporary files

**For Updater Portable profile:**
1. **Publishes** updater application (Portable profile)
2. **Creates** zip archive with updater executable
3. **Saves** to `Release\FlatPatternExporter.Updater-v{VERSION}-x64.zip`
4. **Cleans up** temporary files

**Resulting structure:**
```
Release\
├── FlatPatternExporter-v3.0.0-x64-Deploy.zip              # Deploy profile archive
│   ├── FlatPatternExporter.exe
│   ├── *.dll                                              # All dependencies
│   └── .buildtype                                         # Contains "Deploy"
│
├── FlatPatternExporter-v3.0.0-x64-Portable.zip            # Portable profile archive
│   ├── FlatPatternExporter.exe                            # ~150 MB (SingleFile)
│   └── .buildtype                                         # Contains "Portable"
│
├── FlatPatternExporter-v3.0.0-x64-FrameworkDependent.zip  # Framework-dependent profile archive
│   ├── FlatPatternExporter.exe
│   ├── *.dll                                              # Application libraries only
│   └── .buildtype                                         # Contains "FrameworkDependent"
│
└── FlatPatternExporter.Updater-v3.0.0-x64.zip             # Updater archive
    └── FlatPatternExporter.Updater.exe                    # ~80 MB (SingleFile)
```

**Build Type Detection:**
The `.buildtype` marker file enables automatic update system to download correct archive matching current installation type. The updater is distributed separately and downloaded automatically when needed.

---

## Creating a GitHub Release

`create-release-draft.bat` creates the tag `v{VERSION}` and a GitHub release with the four archives of the current version from `Release\`.

```bash
create-release-draft.bat                                # Draft with a notes template, asks for confirmation
create-release-draft.bat notes.md                       # Draft with notes from a file
create-release-draft.bat notes.md --publish             # Publish immediately
create-release-draft.bat notes.md --publish --yes       # No questions (for scripts)
```

- Requires GitHub CLI (`gh`) and all commits pushed to the remote branch
- Release title: `Flat Pattern Exporter {VersionPrefix}`, tag: `v{VersionPrefix}.{GitCommitCount}`
- `{VERSION}` in the notes file is replaced with the full version
- An existing tag or release is replaced only after confirmation; with `--yes` the script stops instead

Typical release:
```bash
git push
publish.bat all
create-release-draft.bat notes.md --publish
```

### Release notes format

The update window of the installed application shows release notes as plain text, so keep the Markdown simple:
- One sentence about the release, then `##` sections: What's new, Improvements, Fixes, Good to know, Downloads
- Plain `-` lists; no tables, HTML (`<details>`), images or bold text
- Describe what the user gets, not what changed in the code
- End with `Full changelog: https://github.com/isinicyn/FlatPatternExporter/compare/v{PREVIOUS}...v{VERSION}`

---

## Recommendations

### For installers (Inno Setup, WiX, NSIS):
✅ **Use Deploy Profile**
- Transparent file structure
- Easy to manage components
- Can update individual DLLs
- Extract zip archive contents for installer source

### For ZIP archives (Portable version):
✅ **Use Portable Profile**
- Single .exe file
- No installation required
- User-friendly
- Ready to distribute as-is

### For GitHub Releases (automatic updates):
✅ **Upload all build types + updater**
- `FlatPatternExporter-v{VERSION}-x64-Deploy.zip`
- `FlatPatternExporter-v{VERSION}-x64-Portable.zip`
- `FlatPatternExporter-v{VERSION}-x64-FrameworkDependent.zip`
- `FlatPatternExporter.Updater-v{VERSION}-x64.zip`
- Users can choose appropriate build type
- Automatic update system downloads matching archive

### For corporate environments:
✅ **Use FrameworkDependent Profile**
- Minimal size
- Centralized .NET Runtime management
- Requires .NET 8.0 Runtime installation
- Extract zip archive contents for deployment

---

## Publishing from Visual Studio

1. **Open** the project in Visual Studio
2. **Right-click** on the project → **Publish...**
3. **Select** profile:
   - `DeployProfile`
   - `PortableProfile`
   - `FrameworkDependentProfile`
4. **Click** **Publish**

---

## Additional Information

### Target Platform
- **Framework:** .NET 8.0 Windows
- **Runtime:** win-x64
- **OS Version:** Windows 10.0.26100.0+

### Dependencies
- Autodesk Inventor Interop
- netDxf.netstandard (v3.0.1)
- Svg.Skia (v3.2.1)
- ClosedXML (v0.105.0)
- stdole (v17.14.40260)

### Versioning
Application version is generated automatically based on Git:
- Format: `{VersionPrefix}.{GitCommitCount}`, e.g. `3.1.0.705`; `VersionPrefix` is set in `FlatPatternExporter.csproj`
- Informational version contains Git commit hash

---

## Troubleshooting

### Error: "Could not find a part of the path"
- Make sure projects are compiled in Release configuration
- Verify all dependencies are present

### Error: "The framework 'Microsoft.NETCore.App' version '8.0.0' was not found"
- Install .NET 8.0 SDK: https://dotnet.microsoft.com/download/dotnet/8.0

### Error: "git is not recognized"
- Make sure Git is installed and added to PATH
- Versioning depends on Git to generate build number

---

## Contact

For bug reports and suggestions, use Issues in the GitHub repository.
