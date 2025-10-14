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
FlatPatternExporter.Updater-v3.0.0.zip
└── FlatPatternExporter.Updater.exe  # Updater
```

**Process:**
1. Publishes updater (Portable profile)
2. Creates zip archive with updater
3. Saves to `Release\FlatPatternExporter.Updater-v{VERSION}.zip`
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
1. Deploy          - Ready files for installer (zip archive with separate DLLs)
2. Portable        - Portable version (zip archive with single .exe file)
3. Framework       - Depends on .NET 8 Runtime (zip archive, minimal size)
4. Updater Portable - Updater archive only (for deployment)
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
2. **Creates** `.buildtype` marker file (Deploy/Portable/FrameworkDependent)
3. **Creates** zip archive with all files
4. **Saves** to `Release\FlatPatternExporter-v{VERSION}-{BUILD_TYPE}.zip`
5. **Cleans up** temporary files

**For Updater Portable profile:**
1. **Publishes** updater application (Portable profile)
2. **Creates** zip archive with updater executable
3. **Saves** to `Release\FlatPatternExporter.Updater-v{VERSION}.zip`
4. **Cleans up** temporary files

**Resulting structure:**
```
Release\
├── FlatPatternExporter-v3.0.0-Deploy.zip              # Deploy profile archive
│   ├── FlatPatternExporter.exe
│   ├── *.dll                                          # All dependencies
│   └── .buildtype                                     # Contains "Deploy"
│
├── FlatPatternExporter-v3.0.0-Portable.zip            # Portable profile archive
│   ├── FlatPatternExporter.exe                        # ~150 MB (SingleFile)
│   └── .buildtype                                     # Contains "Portable"
│
├── FlatPatternExporter-v3.0.0-FrameworkDependent.zip  # Framework-dependent profile archive
│   ├── FlatPatternExporter.exe
│   ├── *.dll                                          # Application libraries only
│   └── .buildtype                                     # Contains "FrameworkDependent"
│
└── FlatPatternExporter.Updater-v3.0.0.zip             # Updater archive
    └── FlatPatternExporter.Updater.exe                # ~80 MB (SingleFile)
```

**Build Type Detection:**
The `.buildtype` marker file enables automatic update system to download correct archive matching current installation type. The updater is distributed separately and downloaded automatically when needed.

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
- `FlatPatternExporter-v{VERSION}-Deploy.zip`
- `FlatPatternExporter-v{VERSION}-Portable.zip`
- `FlatPatternExporter-v{VERSION}-FrameworkDependent.zip`
- `FlatPatternExporter.Updater-v{VERSION}.zip`
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
