# Autodesk Registry Cleaner v1.0.0

A modular, high-integrity tool for completely removing all Autodesk products from Windows systems. This tool is designed for situations where you need to perform a clean uninstall before reinstalling Autodesk products.

## ⚠️ Important Safety Warning

**THIS TOOL WILL PERMANENTLY DELETE ALL AUTODESK SOFTWARE AND DATA FROM YOUR SYSTEM**

- Back up any important project files before running this tool
- This action cannot be undone
- Only use this tool if you need to completely remove ALL Autodesk products
- Always run as Administrator

## Features

### Modular Architecture
- **Registry Scanner**: Scans and removes Autodesk registry entries from HKLM and HKCU
- **File System Cleaner**: Removes Autodesk files, directories, and temporary data
- **Backup System**: Creates backups of registry and files before removal
- **Safety Validation**: Ensures only Autodesk-related entries are targeted

### Advanced Capabilities
- **Dry Run Mode**: Preview changes without making them
- **Selective Cleaning**: Choose user-only or system-only cleanup
- **Progress Reporting**: Detailed status and error reporting
- **Administrator Validation**: Requires proper privileges for system modifications

## Autodesk Products Supported

This tool removes traces of all Autodesk products including:
- AutoCAD
- Maya
- 3ds Max
- Revit
- Inventor
- Fusion 360
- Vault
- Navisworks
- Mudbox
- MotionBuilder
- Alias Products
- And all other Autodesk software

## System Requirements

- Windows 10/11 (x64)
- .NET 10 Runtime
- Administrator privileges
- Minimum 100MB free disk space (for backups)

## Installation

1. Download the latest release
2. Extract to a folder
3. Right-click `Autodesk_Cleaner.exe` and select "Run as Administrator"

## Usage

### Basic Usage
```cmd
# Run with default settings (includes backup and confirmation)
Autodesk_Cleaner.exe

# Preview changes without making them
Autodesk_Cleaner.exe --dry-run

# Clean without creating backups
Autodesk_Cleaner.exe --no-backup
```

### Advanced Options
```cmd
# Clean only user registry and files
Autodesk_Cleaner.exe --user-only

# Clean only system registry and files
Autodesk_Cleaner.exe --system-only

# Specify custom backup location
Autodesk_Cleaner.exe --backup-path "C:\MyBackups"

# Combine options
Autodesk_Cleaner.exe --dry-run --user-only
```

### Command Line Options

| Option | Short | Description |
|--------|-------|-------------|
| `--dry-run` | `-d` | Preview changes without making them |
| `--no-backup` | `-n` | Skip creating backups |
| `--user-only` | `-u` | Clean only user registry and files |
| `--system-only` | `-s` | Clean only system registry and files |
| `--backup-path PATH` | | Specify custom backup location |

## What Gets Removed

### Registry Entries
- `HKEY_LOCAL_MACHINE\SOFTWARE\Autodesk`
- `HKEY_CURRENT_USER\SOFTWARE\Autodesk`
- Autodesk entries in Windows Uninstall lists
- Autodesk-related installer entries

### Files and Directories
- `C:\Program Files\Autodesk`
- `C:\Program Files\Common Files\Autodesk Shared`
- `C:\Program Files (x86)\Autodesk`
- `C:\Program Files (x86)\Common Files\Autodesk Shared`
- `C:\ProgramData\Autodesk`
- `%USERPROFILE%\Autodesk`
- `%LOCALAPPDATA%\Autodesk`
- `%APPDATA%\Autodesk`
- ADSK files in `C:\ProgramData\FLEXnet`
- Autodesk temporary files

## Safety Features

### Backup System
- Automatic registry backup to `.reg` files
- File backup to compressed `.zip` archives
- Timestamped backup files
- Configurable backup location

### Validation
- Administrator privilege verification
- Autodesk-only pattern matching
- File existence checks before removal
- Read-only attribute handling

### Error Handling
- Comprehensive error reporting
- Graceful failure handling
- Detailed operation summaries
- Progress tracking

## Post-Cleanup Recommendations

After running this tool:

1. **Restart your computer** before installing new Autodesk products
2. **Clear browser cache** if using web-based installers
3. **Temporarily disable antivirus** during Autodesk installation
4. **Run Windows Update** before installing new Autodesk products
5. **Check for Windows .NET Framework** updates if needed

## Technical Architecture

### Core Modules
- `IRegistryScanner` - Registry operations interface
- `AutodeskRegistryScanner` - Registry implementation
- `IFileSystemCleaner` - File system operations interface
- `AutodeskFileSystemCleaner` - File system implementation

### Safety Design
- Immutable data structures
- Async/await for responsive UI
- Resource disposal patterns
- Exception handling at all levels

## Building from Source

### Prerequisites
- Visual Studio 2022 or JetBrains Rider
- .NET 10 SDK
- Windows 10 SDK

### Build Steps
```cmd
# Clone the repository
git clone <repository-url>
cd Autodesk_Cleaner

# Restore packages and build
dotnet restore
dotnet build --configuration Release

# Run tests (if available)
dotnet test

# Publish single-file executable
dotnet publish --configuration Release --runtime win-x64 --self-contained false --output ./publish
```

## Contributing

This project follows strict coding standards:
- C# 14 language features
- .NET 10 target framework
- Modular architecture principles
- Comprehensive error handling
- Full XML documentation

## License

Copyright © 2025 Furr-Tec. All rights reserved.

## Support

For issues related to:
- **Tool functionality**: Create an issue in this repository
- **Autodesk installation**: Contact Autodesk Support
- **Windows registry**: Consult Microsoft documentation

## Disclaimer

This tool is provided "as is" without warranty. Use at your own risk. Always backup important data before using system modification tools. The authors are not responsible for any data loss or system damage.
