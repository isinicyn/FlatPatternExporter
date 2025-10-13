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

### 4. **GitHub Release Profile** (For automatic updates)
Creates a zip archive with both main application and updater for GitHub Releases.

**Characteristics:**
- ✅ Uses Portable profile (SingleFile)
- ✅ Creates zip archive with both executables
- ✅ Archive saved to `Release\GitHub\`
- ✅ Automatic cleanup of temporary files
- 📦 Archive size: ~230 MB
- 🎯 Use case: GitHub Releases for automatic updates

**Archive structure:**
```
FlatPatternExporter-v2.1.0.677.zip
├── FlatPatternExporter.exe          # Main application
└── FlatPatternExporter.Updater.exe  # Updater
```

**Process:**
1. Publishes main application (Portable profile)
2. Publishes updater (Portable profile)
3. Copies both .exe to `Release\GitHub\`
4. Creates zip archive
5. Deletes source .exe files (keeps only zip)

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
1. Deploy          - Ready files for installer (separate DLLs)
2. Portable        - Portable version (single .exe file)
3. Framework       - Depends on .NET 8 Runtime (minimal size)
4. GitHub Release  - Create zip archive for GitHub Release
5. All             - Publish all profiles
6. Exit            - Exit
```

#### Command line mode:
```bash
# Deploy profile
publish.bat deploy

# Portable profile
publish.bat portable

# Framework-dependent profile
publish.bat framework

# All profiles
publish.bat all
```

### What does the script do?

**For Deploy, Portable, Framework profiles:**
1. **Publishes** main application with the selected profile
2. **Copies** results to the `Release\<ProfileName>\` folder
3. **Creates** ready-to-package structure

**For GitHub Release profile:**
1. **Publishes** both main application and updater (Portable profile)
2. **Creates** zip archive with both executables
3. **Saves** to `Release\GitHub\FlatPatternExporter-v{VERSION}.zip`
4. **Cleans up** temporary files

**Resulting structure:**
```
Release\
├── Deploy\                              # Deploy profile
│   ├── FlatPatternExporter.exe
│   └── *.dll                            # All dependencies
│
├── Portable\                            # Portable profile
│   └── FlatPatternExporter.exe          # ~150 MB (everything included, SingleFile)
│
├── FrameworkDependent\                  # Framework-dependent profile
│   ├── FlatPatternExporter.exe
│   └── *.dll                            # Application libraries (without .NET Runtime)
│
└── GitHub\                              # GitHub Release profile
    └── FlatPatternExporter-v2.1.0.677.zip  # Archive with main app + updater
```

**Note:** Updater is only included in the GitHub Release archive. Regular profiles (Deploy, Portable, FrameworkDependent) contain only the main application, as the updater is automatically downloaded from GitHub when needed.

---

## Recommendations

### For installers (Inno Setup, WiX, NSIS):
✅ **Use Deploy Profile**
- Transparent file structure
- Easy to manage components
- Can update individual DLLs

### For ZIP archives (Portable version):
✅ **Use Portable Profile**
- Single .exe file
- No installation required
- User-friendly

### For GitHub Releases (automatic updates):
✅ **Use GitHub Release Profile**
- Zip archive with main app + updater
- Automatic version in filename
- Ready for upload to GitHub Releases
- Supports automatic update system

### For corporate environments:
✅ **Use FrameworkDependent Profile**
- Minimal size
- Centralized .NET Runtime management
- Requires .NET 8.0 Runtime installation

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
- Format: `2.1.0.{GitCommitCount}`
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
