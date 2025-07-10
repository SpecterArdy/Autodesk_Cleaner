# Autodesk Cleaner

Autodesk Cleaner is a command-line utility for cleaning up redundant Autodesk-related files and registry entries on Windows systems. This tool is particularly useful for IT administrators and power users who need to maintain clean systems free from leftover Autodesk components after installations and removals.

## Features

- **Comprehensive Scanning**: Identifies thousands of Autodesk-related file and registry entries.
- **Backup Creation**: Backs up registry entries before deletion to ensure data safety.
- **Service Management**: Stops Autodesk-related services to unlock files and registry keys during cleanup.
- **Process Management**: Identifies and terminates processes locking Autodesk files.
- **Detailed Logging**: Provides verbose logs to track cleanup processes and errors.
- **Dry Run Option**: Allows users to preview changes without making actual modifications.
- **Configurable Paths**: Scans and cleans paths configurable via a settings file.

## Requirements

- Windows OS
- Administrator privileges (application must be run as an admin)
- .NET SDK (compatible with .NET 10)

## Installation

Clone the repository:
```shell
 git clone (https://github.com/SpecterArdy/Autodesk_Cleaner.git)
```
Navigate into the directory:
```shell
 cd Autodesk_Cleaner
```
Build the project:
```shell
 dotnet build
```

## Usage

Run the application with administrator privileges:
```shell
 dotnet run --project Autodesk_Cleaner/Autodesk_Cleaner.csproj
```

### Options

- `--help`: Display help about the command options.
- `--dry-run`: Execute the scan without making any changes.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request for any bug fixes or feature additions.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

