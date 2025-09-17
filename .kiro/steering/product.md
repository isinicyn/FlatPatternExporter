# Product Overview

## FlatPatternExporter (FPExport)

A Windows desktop application for exporting flat patterns from Autodesk Inventor sheet metal parts to DXF format.

### Core Functionality
- Scans Inventor assemblies and parts for sheet metal components
- Exports flat patterns to DXF files with customizable layer settings
- Supports multiple processing methods (assembly structure traversal or BOM-based)
- Provides advanced DXF optimization and spline replacement options
- Integrates directly with Autodesk Inventor through COM interop

### Key Features
- **Multi-format Export**: DXF export with various AutoCAD version compatibility
- **Layer Management**: Customizable layer settings for different geometry types (profiles, bend lines, tool centers, etc.)
- **File Organization**: Flexible output folder structure with material/thickness sorting
- **Batch Processing**: Handles multiple parts in assemblies with progress tracking
- **Conflict Resolution**: Detects and manages naming conflicts during export

### Target Users
- CAD engineers working with Inventor sheet metal designs
- Manufacturing professionals requiring flat pattern data for CNC/laser cutting
- Design teams needing automated DXF export workflows

### Technical Context
- Requires Autodesk Inventor installation for COM interop functionality
- Processes .ipt (part) and .iam (assembly) files
- Outputs industry-standard DXF files for manufacturing workflows