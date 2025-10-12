# FlatPatternExporter Publishing Guide

This document describes the application publishing process for various deployment scenarios.

## Publish Profiles

### 1. **Deploy Profile** (Recommended for installers)
Publishes the application with separate DLL files.

**Main Application:**
- File: `FlatPatternExporter\Properties\PublishProfiles\DeployProfile.pubxml`
- Output: `FlatPatternExporter\bin\publish\deploy\`

**Updater:**
- File: `FlatPatternExporter.Updater\Properties\PublishProfiles\DeployProfile.pubxml`
- Output: `FlatPatternExporter.Updater\bin\publish\deploy\`

**Characteristics:**
- ✅ Self-contained (.NET Runtime included)
- ✅ Separate DLL files
- ✅ PublishReadyToRun (startup optimization)
- ✅ Transparent file structure
- 📦 Size: ~100-120 MB
- 🎯 Use case: Inno Setup, WiX, NSIS installers

**Manual publishing:**
```bash
# Main application
dotnet publish FlatPatternExporter\FlatPatternExporter.csproj --configuration Release /p:PublishProfile=DeployProfile

# Updater
dotnet publish FlatPatternExporter.Updater\FlatPatternExporter.Updater.csproj --configuration Release /p:PublishProfile=DeployProfile
```

---

### 2. **Portable Profile** (For ZIP archives)
Publishes the application as a single executable file.

**Main Application:**
- File: `FlatPatternExporter\Properties\PublishProfiles\PortableProfile.pubxml`
- Output: `FlatPatternExporter\bin\publish\portable\`

**Updater:**
- File: `FlatPatternExporter.Updater\Properties\PublishProfiles\PortableProfile.pubxml`
- Output: `FlatPatternExporter.Updater\bin\publish\portable\`

**Characteristics:**
- ✅ Self-contained (.NET Runtime included)
- ✅ PublishSingleFile (single .exe)
- ✅ EnableCompressionInSingleFile
- ✅ IncludeNativeLibrariesForSelfExtract
- 📦 Size: ~150+ MB
- 🎯 Use case: Portable version, ZIP archives

**Manual publishing:**
```bash
# Main application
dotnet publish FlatPatternExporter\FlatPatternExporter.csproj --configuration Release /p:PublishProfile=PortableProfile

# Updater
dotnet publish FlatPatternExporter.Updater\FlatPatternExporter.Updater.csproj --configuration Release /p:PublishProfile=PortableProfile
```

---

### 3. **FrameworkDependent Profile** (Minimal size)
Publishes the application without including .NET Runtime.

**Main Application:**
- File: `FlatPatternExporter\Properties\PublishProfiles\FrameworkDependentProfile.pubxml`
- Output: `FlatPatternExporter\bin\publish\framework-dependent\`

**Updater:**
- File: `FlatPatternExporter.Updater\Properties\PublishProfiles\FrameworkDependentProfile.pubxml`
- Output: `FlatPatternExporter.Updater\bin\publish\framework-dependent\`

**Characteristics:**
- ⚠️ Requires .NET 8.0 Runtime on target machine
- ✅ Minimal size
- ✅ Separate DLL files
- 📦 Size: ~5-10 MB
- 🎯 Use case: Corporate environments with centralized .NET Runtime management

**Manual publishing:**
```bash
# Main application
dotnet publish FlatPatternExporter\FlatPatternExporter.csproj --configuration Release /p:PublishProfile=FrameworkDependentProfile

# Updater
dotnet publish FlatPatternExporter.Updater\FlatPatternExporter.Updater.csproj --configuration Release /p:PublishProfile=FrameworkDependentProfile
```

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
4. All             - Publish all profiles
5. Exit            - Exit
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

1. **Publishes** both projects (FlatPatternExporter + Updater) with the selected profile
2. **Copies** results to the `Release\<ProfileName>\` folder
3. **Creates** ready-to-package structure

**Resulting structure:**
```
Release\
├── Deploy\                              # Deploy profile
│   ├── FlatPatternExporter.exe
│   ├── FlatPatternExporter.Updater.exe
│   └── *.dll                            # All dependencies of both applications
│
├── Portable\                            # Portable profile
│   ├── FlatPatternExporter.exe          # ~150 MB (everything included, SingleFile)
│   └── FlatPatternExporter.Updater.exe  # ~80 MB (everything included, SingleFile)
│
└── FrameworkDependent\                  # Framework-dependent profile
    ├── FlatPatternExporter.exe
    ├── FlatPatternExporter.Updater.exe
    └── *.dll                            # Application libraries (without .NET Runtime)
```

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
