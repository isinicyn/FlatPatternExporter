# Project Structure

## Solution Organization
```
FlatPatternExporter.sln          # Visual Studio solution file
├── FlatPatternExporter/         # Main WPF application project
├── HELP/                        # Documentation and help files
├── README.md                    # Project documentation
├── LICENSE.txt                  # License information
└── REF.markdown                 # Refactoring guidelines
```

## Main Project Structure (`FlatPatternExporter/`)

### Core Application Files
```
├── App.xaml                     # Application entry point and global resources
├── App.xaml.cs                  # Application startup logic
├── AssemblyInfo.cs              # Assembly metadata
├── FlatPatternExporter.csproj   # Project configuration
└── FPExport.ico                 # Application icon
```

### Windows and Views
```
├── FlatPatternExporterMainWindow.xaml      # Main application window
├── FlatPatternExporterMainWindow.xaml.cs   # Main window code-behind
├── FlatPatternExporterMainWindow.xaml.part2.cs  # Extended code-behind
├── AboutWindow.xaml             # About dialog
├── AboutWindow.xaml.cs          # About dialog code-behind
├── ConflictDetailsWindow.xaml   # Conflict resolution dialog
├── ConflictDetailsWindow.xaml.cs
├── SelectIPropertyWindow.xaml   # Property selection dialog
└── SelectIPropertyWindow.xaml.cs
```

### Custom Controls and Components
```
├── Components/                  # Reusable UI components (empty - ready for expansion)
├── LayerSettingControl.xaml     # Layer configuration control
├── LayerSettingControl.xaml.cs
├── TextWithFxIndicator.xaml     # Text control with effects
└── TextWithFxIndicator.xaml.cs
```

### MVVM Architecture Folders
```
├── ViewModels/                  # View models (empty - ready for MVVM implementation)
├── Models/                      # Data models (empty - ready for expansion)
├── Commands/                    # ICommand implementations (empty - ready for commands)
├── Services/                    # Business logic services (empty - ready for services)
└── Converters/                  # Value converters (empty - ready for converters)
```

### Core Business Logic
```
├── DxfGenerator.cs              # DXF file generation and thumbnail creation
├── PropertyManager.cs           # Inventor property management
├── SettingsManager.cs           # Application settings persistence
├── TokenService.cs              # File naming token management
├── LayerSettingsClasses.cs      # Layer configuration data classes
├── MarshalCore.cs               # COM interop marshalling utilities
└── IPictureDispConverter.cs     # Image conversion utilities
```

### UI Resources and Styling
```
├── Styles/                      # XAML style resources
│   ├── ColorResources.xaml      # Color palette and brushes
│   ├── IconResources.xaml       # SVG icons and geometry
│   ├── GeneralStyles.xaml       # General UI styles
│   └── DataGridStyles.xaml      # Data grid specific styles
```

### Build Artifacts (Generated)
```
├── bin/                         # Compiled binaries
├── obj/                         # Build intermediate files
└── Properties/                  # Assembly properties
    └── launchSettings.json      # Debug launch configuration
```

## Naming Conventions

### Files and Folders
- **PascalCase** for all C# files and classes
- **Descriptive names** reflecting functionality (e.g., `LayerSettingControl`, `ConflictDetailsWindow`)
- **Grouped by purpose** (Windows, Controls, Services, etc.)

### XAML Resources
- **Semantic naming** for styles (e.g., `BaseButtonStyle`, `DocumentInfoLabelStyle`)
- **Categorized resource files** (Colors, Icons, General styles)
- **Consistent resource keys** following pattern: `{Purpose}{Element}Style`

### Code Organization
- **Partial classes** for large windows (`.xaml.part2.cs` pattern)
- **Separation of concerns** with dedicated folders for each layer
- **Empty folders** prepared for future MVVM implementation

## Architecture Notes

### Current State
- **Code-behind heavy** - Most logic in `.xaml.cs` files
- **Monolithic windows** - Large main window with embedded logic
- **Direct COM interop** - Inventor integration throughout UI layer

### Recommended Evolution
- **MVVM migration** - Move logic to ViewModels and Services
- **Command pattern** - Replace event handlers with ICommand
- **Dependency injection** - Service registration and resolution
- **Unit testing** - Testable business logic separation

### Extension Points
- `Commands/` - Ready for RelayCommand implementations
- `Services/` - Business logic extraction target
- `ViewModels/` - Data binding and UI state management
- `Models/` - Domain objects and data transfer objects
- `Converters/` - UI data transformation logic