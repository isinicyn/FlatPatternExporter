# Technology Stack

## Framework & Runtime
- **.NET 8.0** with Windows-specific targeting (`net8.0-windows10.0.26100.0`)
- **WPF (Windows Presentation Foundation)** for desktop UI
- **Windows Forms** integration enabled for legacy components
- **C# 12** with nullable reference types and implicit usings enabled

## Key Dependencies
- **Autodesk.Inventor.Interop** - COM interop for Inventor integration
- **netDxf.netstandard** (v3.0.1) - DXF file generation and manipulation
- **stdole** (v17.14.40260) - Standard OLE types for COM interop

## Architecture Patterns
- **MVVM (Model-View-ViewModel)** - WPF application architecture
- **Command Pattern** - UI interactions through ICommand implementations
- **Service Layer** - Business logic separation in Services folder
- **Component-Based UI** - Reusable UserControls and custom components

## Build System
- **MSBuild** with modern SDK-style project format
- **Git-based versioning** - Automatic version generation from git commit count and hash
- **Debug/Release configurations** with optimized settings

## Common Build Commands

### Development
```bash
# Build the solution
dotnet build FlatPatternExporter.sln

# Build in Release mode
dotnet build FlatPatternExporter.sln -c Release

# Run the application
dotnet run --project FlatPatternExporter

# Clean build artifacts
dotnet clean FlatPatternExporter.sln
```

### Publishing
```bash
# Publish self-contained executable
dotnet publish FlatPatternExporter -c Release -r win-x64 --self-contained

# Publish framework-dependent
dotnet publish FlatPatternExporter -c Release -r win-x64 --no-self-contained
```

### Testing & Analysis
```bash
# Restore NuGet packages
dotnet restore

# Check for outdated packages
dotnet list package --outdated

# Format code
dotnet format
```

## Development Environment
- **Visual Studio 2022** (v17.10+) recommended
- **Windows 10/11** required for WPF and Inventor COM interop
- **Autodesk Inventor 2026** for full functionality testing

## Code Style & Standards
- **Microsoft C# Coding Conventions**
- **Modern C# 12 syntax** (primary constructors, collection expressions, raw string literals)
- **XML documentation** for public APIs
- **Nullable reference types** enabled project-wide
- **Russian language** for UI text and user-facing messages