using Autodesk_Cleaner.Core;
using System.Security.Principal;
using NLog;
using Spectre.Console;

namespace Autodesk_Cleaner;

/// csummarye
/// Main entry point for the Autodesk Registry Cleaner console application.
/// c/summarye
internal static class Program
{
    private static Logger? _logger;
    private static ScannerConfig _currentConfig = ScannerConfig.Default;

    /// <summary>
    /// Main entry point for the application.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <returns>Exit code (0 for success, non-zero for error).</returns>
    private static async Task<int> Main(string[] args)
    {
        // Check for debug mode
        bool debugMode = args.Contains("--debug") || args.Contains("-d");
        
        if (debugMode)
        {
            return await RunDebugModeAsync();
        }
        
        try
        {
            // Initialize logging first
            LoggingConfiguration.InitializeLogging();
            _logger = LogManager.GetCurrentClassLogger();
            
            AnsiConsole.MarkupLine("[bold blue]Autodesk Registry Cleaner v1.0.0[/]");
            
            _logger.Info("Application started with interactive menu");
            
            // Check for administrator privileges
            if (!IsRunningAsAdministrator())
            {
                _logger.Error("Application not running as Administrator");
                
                var errorPanel = new Panel(new Markup(
                    "[bold red]ADMINISTRATOR PRIVILEGES REQUIRED[/]\n\n" +
                    "This application must be run as Administrator to modify registry and system files.\n" +
                    "Please restart the application with administrator privileges."))
                {
                    Header = new PanelHeader(" [bold red]ACCESS DENIED[/] "),
                    Border = BoxBorder.Double,
                    BorderStyle = new Style(Color.Red)
                };
                
                AnsiConsole.Write(errorPanel);
                AnsiConsole.MarkupLine("\n[dim]Press any key to exit...[/]");
                Console.ReadKey();
                return 1;
            }
            
            _logger.Info("Administrator privileges confirmed");

            using var menu = new InteractiveMenu();
            
            while (true)
            {
                var option = menu.DisplayMainMenu();
                
                if (option == InteractiveMenu.MenuOption.Exit)
                {
                    _logger.Info("Exiting application");
                    break;
                }
                
                var exitCode = await HandleMenuOptionAsync(option);
                if (exitCode != 0)
                {
                    return exitCode;
                }
            }

            return 0;
        }
        catch (Exception ex)
        {
            _logger?.Fatal(ex, "Critical error in main application");
            
            var errorPanel = new Panel(new Markup(
                $"[bold red]CRITICAL ERROR[/]\n\n" +
                $"An unexpected error occurred: [yellow]{ex.Message}[/]\n\n" +
                "[dim]Check the logs for more detailed information.[/]"))
            {
                Header = new PanelHeader(" [bold red]FATAL ERROR[/] "),
                Border = BoxBorder.Double,
                BorderStyle = new Style(Color.Red)
            };
            
            AnsiConsole.Write(errorPanel);
            
            if (_logger != null)
            {
                AnsiConsole.MarkupLine("[dim]Full stack trace has been logged for debugging.[/]");
            }
            
            AnsiConsole.MarkupLine("\n[dim]Press any key to exit...[/]");
            Console.ReadKey();
            
            return -1;
        }
        finally
        {
            _logger?.Info("Application shutdown");
            LogManager.Shutdown();
        }
    }

    /// <summary>
    /// Handles the selected menu option.
    /// </summary>
    /// <param name="option">The selected menu option.</param>
    /// <returns>Exit code (0 for continue, non-zero for exit).</returns>
    private static async Task<int> HandleMenuOptionAsync(InteractiveMenu.MenuOption option)
    {
        try
        {
            switch (option)
            {
                case InteractiveMenu.MenuOption.ScanRegistryOnly:
                    await HandleScanRegistryAsync();
                    break;
                case InteractiveMenu.MenuOption.ScanFileSystemOnly:
                    await HandleScanFileSystemAsync();
                    break;
                case InteractiveMenu.MenuOption.ScanBoth:
                    await HandleScanBothAsync();
                    break;
                case InteractiveMenu.MenuOption.DryRunCleanup:
                    await HandleDryRunCleanupAsync();
                    break;
                case InteractiveMenu.MenuOption.ActualCleanup:
                    await HandleActualCleanupAsync();
                    break;
                case InteractiveMenu.MenuOption.ConfigureBackup:
                    HandleConfigureBackup();
                    break;
                case InteractiveMenu.MenuOption.ViewConfiguration:
                    HandleViewConfiguration();
                    break;
                case InteractiveMenu.MenuOption.EmergencyAbort:
                    _logger?.Warn("Emergency abort triggered by user");
                    AnsiConsole.MarkupLine("[bold red]Emergency abort initiated. Exiting immediately.[/]");
                    return -1;
                default:
                    _logger?.Warn("Unknown menu option selected: {Option}", option);
                    AnsiConsole.MarkupLine("[red]Unknown option selected.[/]");
                    break;
            }
            
            // Pause after each operation
            AnsiConsole.MarkupLine("\n[dim]Press any key to return to main menu...[/]");
            Console.ReadKey(true);
            
            return 0;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "Error handling menu option: {Option}", option);
            AnsiConsole.WriteException(ex);
            return 0; // Return to menu instead of exiting
        }
    }
    
    /// <summary>
    /// Handles scanning registry only.
    /// </summary>
    private static async Task HandleScanRegistryAsync()
    {
        var config = _currentConfig with { DryRun = true };
        
        await AnsiConsole.Status()
            .StartAsync("[green]Scanning registry for Autodesk entries...[/]", async ctx =>
            {
                using var scanner = new AutodeskRegistryScanner(config);
                var entries = await scanner.ScanRegistryAsync();
                
                ctx.Status("Creating registry report...");
                DisplayRegistryResults(entries);
            });
    }
    
    /// <summary>
    /// Handles scanning file system only.
    /// </summary>
    private static async Task HandleScanFileSystemAsync()
    {
        var config = _currentConfig with { DryRun = true };
        
        await AnsiConsole.Status()
            .StartAsync("[green]Scanning file system for Autodesk files...[/]", async ctx =>
            {
                using var cleaner = new AutodeskFileSystemCleaner(config);
                var entries = await cleaner.ScanFileSystemAsync();
                
                ctx.Status("Creating file system report...");
                DisplayFileSystemResults(entries);
            });
    }
    
    /// <summary>
    /// Handles scanning both registry and file system.
    /// </summary>
    private static async Task HandleScanBothAsync()
    {
        var config = _currentConfig with { DryRun = true };
        
        await AnsiConsole.Status()
            .StartAsync("[green]Performing comprehensive scan...[/]", async ctx =>
            {
                ctx.Status("Scanning registry...");
                using var registryScanner = new AutodeskRegistryScanner(config);
                var registryEntries = await registryScanner.ScanRegistryAsync();
                
                ctx.Status("Scanning file system...");
                using var fileSystemCleaner = new AutodeskFileSystemCleaner(config);
                var fileSystemEntries = await fileSystemCleaner.ScanFileSystemAsync();
                
                ctx.Status("Generating comprehensive report...");
                
                AnsiConsole.Write(new Rule("[bold blue]COMPREHENSIVE SCAN RESULTS[/]"));
                DisplayRegistryResults(registryEntries);
                DisplayFileSystemResults(fileSystemEntries);
            });
    }
    
    /// <summary>
    /// Handles dry run cleanup operation.
    /// </summary>
    private static async Task HandleDryRunCleanupAsync()
    {
        var config = _currentConfig with { DryRun = true };
        await ExecuteCleaningOperationAsync(config);
    }
    
    /// <summary>
    /// Handles actual cleanup operation.
    /// </summary>
    private static async Task HandleActualCleanupAsync()
    {
        var config = _currentConfig with { DryRun = false };
        await ExecuteCleaningOperationAsync(config);
    }
    
    /// <summary>
    /// Handles backup configuration.
    /// </summary>
    private static void HandleConfigureBackup()
    {
        var currentConfig = _currentConfig;
        
        // Display current configuration
        DisplayCurrentBackupConfig(currentConfig);
        
        // Interactive configuration editing
        var createBackup = AnsiConsole.Confirm(
            "Do you want to create backups before cleanup operations?", 
            currentConfig.CreateBackup);
        
        string backupPath = currentConfig.BackupPath;
        if (createBackup)
        {
            backupPath = AnsiConsole.Ask(
                "Enter backup directory path:", 
                currentConfig.BackupPath);
            
            // Validate and create backup directory
            try
            {
                if (!Directory.Exists(backupPath))
                {
                    var createDir = AnsiConsole.Confirm(
                        $"Directory '{backupPath}' does not exist. Create it?");
                    
                    if (createDir)
                    {
                        Directory.CreateDirectory(backupPath);
                        AnsiConsole.MarkupLine($"[green]Created backup directory: {backupPath}[/]");
                    }
                    else
                    {
                        AnsiConsole.MarkupLine("[yellow]Backup path not created. Using default.[/]");
                        backupPath = currentConfig.BackupPath;
                    }
                }
                else
                {
                    AnsiConsole.MarkupLine($"[green]Backup directory exists: {backupPath}[/]");
                }
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error creating backup directory: {ex.Message}[/]");
                AnsiConsole.MarkupLine("[yellow]Using default backup path.[/]");
                backupPath = currentConfig.BackupPath;
            }
        }
        
        var maxDepth = AnsiConsole.Ask(
            "Maximum directory scan depth:", 
            currentConfig.MaxDepth);
        
        // Create new configuration
        var newConfig = currentConfig with 
        {
            CreateBackup = createBackup,
            BackupPath = backupPath,
            MaxDepth = maxDepth
        };
        
        // Store the configuration for future use (in a real implementation, this would be persisted)
        _currentConfig = newConfig;
        
        AnsiConsole.Write(new Rule("[bold green]CONFIGURATION UPDATED[/]"));
        DisplayCurrentBackupConfig(newConfig);
        
        AnsiConsole.MarkupLine("[bold green]Backup configuration updated successfully[/]!");
        AnsiConsole.MarkupLine("[dim]Note: Configuration will be used for subsequent operations in this session.[/]");
    }
    
    /// <summary>
    /// Displays the current backup configuration.
    /// </summary>
    /// <param name="config">The configuration to display.</param>
    private static void DisplayCurrentBackupConfig(ScannerConfig config)
    {
        var table = new Table();
        table.AddColumn("[bold]Setting[/]");
        table.AddColumn("[bold]Current Value[/]");
        table.AddColumn("[bold]Description[/]");
        
        table.AddRow("Create Backup", config.CreateBackup ? "[green]Yes[/]" : "[red]No[/]", "Create registry/file backups before deletion");
        table.AddRow("Backup Path", $"[yellow]{config.BackupPath}[/]", "Location where backups are stored");
        table.AddRow("Max Depth", config.MaxDepth.ToString(), "Maximum directory depth for file scanning");
        table.AddRow("Include User Registry", config.IncludeUserHive ? "[green]Yes[/]" : "[red]No[/]", "Scan user registry (HKEY_CURRENT_USER)");
        table.AddRow("Include System Registry", config.IncludeLocalMachine ? "[green]Yes[/]" : "[red]No[/]", "Scan system registry (HKEY_LOCAL_MACHINE)");
        
        var panel = new Panel(table)
        {
            Header = new PanelHeader(" [bold blue]CURRENT BACKUP CONFIGURATION[/] "),
            Border = BoxBorder.Rounded
        };
        
        AnsiConsole.Write(panel);
    }
    
    /// <summary>
    /// Handles viewing current configuration.
    /// </summary>
    private static void HandleViewConfiguration()
    {
        var config = _currentConfig;
        
        var table = new Table()
        {
            Border = TableBorder.Rounded,
            BorderStyle = new Style(Color.Blue)
        };
        
        table.AddColumn("[bold]Configuration Item[/]");
        table.AddColumn("[bold]Value[/]");
        table.AddColumn("[bold]Description[/]");
        
        table.AddRow("Dry Run Mode", config.DryRun ? "[green]Enabled[/]" : "[red]Disabled[/]", "Preview changes without making them");
        table.AddRow("Create Backup", config.CreateBackup ? "[green]Enabled[/]" : "[red]Disabled[/]", "Create backups before deletion");
        table.AddRow("Backup Path", $"[yellow]{config.BackupPath}[/]", "Backup storage location");
        table.AddRow("Include User Registry", config.IncludeUserHive ? "[green]Yes[/]" : "[red]No[/]", "Scan HKEY_CURRENT_USER");
        table.AddRow("Include System Registry", config.IncludeLocalMachine ? "[green]Yes[/]" : "[red]No[/]", "Scan HKEY_LOCAL_MACHINE");
        table.AddRow("Max Scan Depth", config.MaxDepth.ToString(), "Maximum directory depth");
        
        var panel = new Panel(table)
        {
            Header = new PanelHeader(" [bold blue]CURRENT CONFIGURATION[/] "),
            Border = BoxBorder.Double
        };
        
        AnsiConsole.Write(panel);
    }
    
    /// <summary>
    /// Displays registry scan results.
    /// </summary>
    /// <param name="entries">Registry entries found.</param>
    private static void DisplayRegistryResults(IReadOnlyCollection<RegistryEntry> entries)
    {
        var table = new Table()
        {
            Border = TableBorder.Rounded,
            BorderStyle = new Style(Color.Green)
        };
        
        table.AddColumn("[bold]Registry Key[/]");
        table.AddColumn("[bold]Display Name[/]");
        table.AddColumn("[bold]Type[/]");
        
        foreach (var entry in entries.Take(20)) // Show first 20
        {
            table.AddRow(
                $"[dim]{entry.KeyPath}[/]",
                entry.DisplayName ?? "[dim]N/A[/]",
                $"[yellow]{entry.EntryType}[/]"
            );
        }
        
        if (entries.Count > 20)
        {
            table.AddRow($"[dim]... and {entries.Count - 20} more entries[/]", "", "");
        }
        
        var panel = new Panel(table)
        {
            Header = new PanelHeader($" [bold green]REGISTRY SCAN RESULTS ({entries.Count} entries)[/] ")
        };
        
        AnsiConsole.Write(panel);
    }
    
    /// <summary>
    /// Displays file system scan results.
    /// </summary>
    /// <param name="entries">File system entries found.</param>
    private static void DisplayFileSystemResults(IReadOnlyCollection<FileSystemEntry> entries)
    {
        var table = new Table()
        {
            Border = TableBorder.Rounded,
            BorderStyle = new Style(Color.Blue)
        };
        
        table.AddColumn("[bold]Path[/]");
        table.AddColumn("[bold]Type[/]");
        table.AddColumn("[bold]Size[/]");
        
        foreach (var entry in entries.Take(20)) // Show first 20
        {
            var sizeText = entry.EntryType == FileSystemEntryType.File 
                ? FormatFileSize(entry.Size) 
                : "[dim]Directory[/]";
                
            table.AddRow(
                $"[dim]{entry.Path}[/]",
                $"[cyan]{entry.EntryType}[/]",
                sizeText
            );
        }
        
        if (entries.Count > 20)
        {
            table.AddRow($"[dim]... and {entries.Count - 20} more entries[/]", "", "");
        }
        
        var panel = new Panel(table)
        {
            Header = new PanelHeader($" [bold blue]FILE SYSTEM SCAN RESULTS ({entries.Count} entries)[/] ")
        };
        
        AnsiConsole.Write(panel);
    }
    
    /// <summary>
    /// Executes the main cleaning operation.
    /// </summary>
    /// <param name="config">The scanner configuration.</param>
    /// <returns>Exit code.</returns>
    private static async Task<int> ExecuteCleaningOperationAsync(ScannerConfig config)
    {
        var totalErrors = new List<string>();
        var overallSuccess = true;

        try
        {
            AnsiConsole.Write(new Rule("[bold yellow]AUTODESK CLEANUP OPERATION[/]"));
            
            if (config.DryRun)
            {
                AnsiConsole.MarkupLine("[bold yellow]DRY RUN MODE - No actual changes will be made[/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[bold red]LIVE MODE - Changes will be permanent[/]");
            }

            // Phase 1: Registry Cleanup
            AnsiConsole.Write(new Rule("[bold blue]Phase 1: Registry Cleanup[/]"));
            
            AnsiConsole.MarkupLine("[green]Scanning registry for Autodesk entries...[/]");
            IReadOnlyCollection<RegistryEntry> registryEntries;
            using (var registryScanner = new AutodeskRegistryScanner(config))
            {
                registryEntries = await registryScanner.ScanRegistryAsync();
            }
            AnsiConsole.MarkupLine($"[green]Registry scan completed. Found {registryEntries.Count} raw entries.[/]");

            // Check how many entries are actually valid for cleaning
            var validRegistryCount = 0;
            if (registryEntries.Count > 0)
            {
                using var registryValidator = new AutodeskRegistryScanner(config);
                validRegistryCount = registryValidator.GetValidEntryCount(registryEntries);
            }
                
            AnsiConsole.MarkupLine($"[green]Found {validRegistryCount} registry entries to clean.[/]");

            if (validRegistryCount > 0)
            {
                DisplayRegistryResults(registryEntries);
                
                var registryResult = await AnsiConsole.Status()
                    .StartAsync("[yellow]Processing registry entries...[/]", async ctx =>
                    {
                        using var registryScanner = new AutodeskRegistryScanner(config);
                        return await registryScanner.RemoveEntriesAsync(registryEntries);
                    });
                    
                DisplayRemovalResult("Registry", registryResult);
                
                if (!registryResult.IsSuccessful)
                {
                    overallSuccess = false;
                    totalErrors.AddRange(registryResult.Errors);
                }
            }
            else
            {
                AnsiConsole.MarkupLine("[dim]No registry entries found to clean.[/]");
            }

            // Phase 2: File System Cleanup
            AnsiConsole.Write(new Rule("[bold blue]Phase 2: File System Cleanup[/]"));
            
            AnsiConsole.MarkupLine("[green]Scanning file system for Autodesk files and directories...[/]");
            IReadOnlyCollection<FileSystemEntry> fileSystemEntries;
            using (var fileSystemCleaner = new AutodeskFileSystemCleaner(config))
            {
                fileSystemEntries = await fileSystemCleaner.ScanFileSystemAsync();
            }
            AnsiConsole.MarkupLine($"[green]File system scan completed. Found {fileSystemEntries.Count} raw entries.[/]");

            // Check how many entries are actually valid for cleaning
            var validFileSystemCount = 0;
            if (fileSystemEntries.Count > 0)
            {
                using var fileSystemValidator = new AutodeskFileSystemCleaner(config);
                validFileSystemCount = fileSystemValidator.GetValidEntryCount(fileSystemEntries);
            }
                
            AnsiConsole.MarkupLine($"[green]Found {validFileSystemCount} file system entries to clean.[/]");

            if (validFileSystemCount > 0)
            {
                DisplayFileSystemResults(fileSystemEntries);
                
                var fileSystemResult = await AnsiConsole.Status()
                    .StartAsync("[yellow]Processing file system entries...[/]", async ctx =>
                    {
                        using var fileSystemCleaner = new AutodeskFileSystemCleaner(config);
                        return await fileSystemCleaner.RemoveEntriesAsync(fileSystemEntries);
                    });
                    
                DisplayRemovalResult("File System", fileSystemResult);
                
                if (!fileSystemResult.IsSuccessful)
                {
                    overallSuccess = false;
                    totalErrors.AddRange(fileSystemResult.Errors);
                }
            }
            else
            {
                AnsiConsole.MarkupLine("[dim]No file system entries found to clean.[/]");
            }

            // Summary
            WriteSummary(overallSuccess, totalErrors, config.DryRun);

            return overallSuccess ? 0 : 1;
        }
        catch (Exception ex)
        {
            AnsiConsole.WriteException(ex);
            return -1;
        }
    }

    /// <summary>
    /// Displays the removal result for a specific phase.
    /// </summary>
    /// <param name="phaseName">The name of the phase.</param>
    /// <param name="result">The removal result.</param>
    private static void DisplayRemovalResult(string phaseName, RemovalResult result)
    {
        var table = new Table()
        {
            Border = TableBorder.Rounded,
            BorderStyle = result.IsSuccessful ? new Style(Color.Green) : new Style(Color.Yellow)
        };
        
        table.AddColumn("[bold]Metric[/]");
        table.AddColumn("[bold]Value[/]");
        
        table.AddRow("Total Entries", result.TotalEntries.ToString());
        table.AddRow("Successfully Removed", $"[green]{result.SuccessfulRemovals}[/]");
        table.AddRow("Failed Removals", result.FailedRemovals > 0 ? $"[red]{result.FailedRemovals}[/]" : "0");
        table.AddRow("Success Rate", $"{result.SuccessRate:F1}%");
        table.AddRow("Duration", $"{result.Duration.TotalSeconds:F2} seconds");
        
        var statusIcon = result.IsSuccessful ? "OK" : "WARNING";
        var statusColor = result.IsSuccessful ? "green" : "yellow";
        
        var panel = new Panel(table)
        {
            Header = new PanelHeader($" [bold {statusColor}]{statusIcon} {phaseName} Cleanup Results[/] ")
        };
        
        AnsiConsole.Write(panel);
        
        if (result.Errors.Count > 0)
        {
            var errorPanel = new Panel(
                string.Join("\n", result.Errors.Take(5).Select(error => $"[red]• {error}[/]")) +
                (result.Errors.Count > 5 ? $"\n[dim]... and {result.Errors.Count - 5} more errors[/]" : ""))
            {
                Header = new PanelHeader($" [bold red]Errors Encountered ({result.Errors.Count})[/] "),
                Border = BoxBorder.Rounded,
                BorderStyle = new Style(Color.Red)
            };
            
            AnsiConsole.Write(errorPanel);
        }
    }

    /// <summary>
    /// Displays the final summary of the operation.
    /// </summary>
    /// <param name="overallSuccess">Whether the overall operation was successful.</param>
    /// <param name="totalErrors">All errors encountered during the operation.</param>
    /// <param name="isDryRun">Whether this was a dry run.</param>
    private static void WriteSummary(bool overallSuccess, List<string> totalErrors, bool isDryRun)
    {
        AnsiConsole.Write(new Rule("[bold blue]OPERATION SUMMARY[/]"));
        
        Panel summaryPanel;
        
        if (isDryRun)
        {
            summaryPanel = new Panel(new Markup(
                "[bold blue]DRY RUN COMPLETED[/]\n\n" +
                "No actual changes were made to your system.\n" +
                "This was a preview of what would be removed in a real cleanup operation."))
            {
                Header = new PanelHeader(" [bold blue]DRY RUN SUMMARY[/] "),
                Border = BoxBorder.Rounded,
                BorderStyle = new Style(Color.Blue)
            };
        }
        else if (overallSuccess)
        {
            summaryPanel = new Panel(new Markup(
                "[bold green]Autodesk cleanup completed successfully[/]!\n\n" +
                "All Autodesk registry entries and files have been removed from your system.\n" +
                "You can now proceed with a fresh installation of Autodesk products."))
            {
                Header = new PanelHeader(" [bold green]SUCCESS[/] "),
                Border = BoxBorder.Double,
                BorderStyle = new Style(Color.Green)
            };
        }
        else
        {
            summaryPanel = new Panel(new Markup(
                $"[bold yellow]Autodesk cleanup completed with some errors.[/]\n\n" +
                $"Total errors encountered: [red]{totalErrors.Count}[/]\n" +
                "Some Autodesk entries may still remain on your system.\n" +
                "You may need to manually remove remaining entries or run the tool again."))
            {
                Header = new PanelHeader(" [bold yellow]PARTIAL SUCCESS[/] "),
                Border = BoxBorder.Rounded,
                BorderStyle = new Style(Color.Yellow)
            };
        }
        
        AnsiConsole.Write(summaryPanel);
        
        // Important notes
        var notesList = new List<string>
        {
            "Restart your computer before installing new Autodesk products",
            "Clear your browser cache if using web-based installers",
            "Temporarily disable antivirus during Autodesk installation",
            "Run Windows Update before installing new Autodesk products"
        };
        
        var notesPanel = new Panel(
            string.Join("\n", notesList.Select(note => $"[yellow]•[/] {note}")))
        {
            Header = new PanelHeader(" [bold cyan]IMPORTANT NOTES[/] "),
            Border = BoxBorder.Rounded,
            BorderStyle = new Style(Color.Blue)
        };
        
        AnsiConsole.Write(notesPanel);
    }


    /// <summary>
    /// Checks if the application is running with administrator privileges.
    /// </summary>
    /// <returns>True if running as administrator, false otherwise.</returns>
    private static bool IsRunningAsAdministrator()
    {
        try
        {
            using var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Runs the application in debug mode without Spectre.Console.
    /// </summary>
    /// <returns>Exit code.</returns>
#pragma warning disable Spectre1000 // Use AnsiConsole instead of System.Console
    private static async Task<int> RunDebugModeAsync()
    {
        Console.WriteLine("=== AUTODESK CLEANER DEBUG MODE ===");
        Console.WriteLine($"Current Directory: {Environment.CurrentDirectory}");
        Console.WriteLine($"Application Directory: {AppDomain.CurrentDomain.BaseDirectory}");
        
        Console.WriteLine("\n=== INITIALIZING LOGGING ===");
        
        // Initialize logging with debug output
        LoggingConfiguration.InitializeLogging();
        var logger = LogManager.GetCurrentClassLogger();
        
        Console.WriteLine("\n=== TESTING LOGGING ===");
        logger.Info("DEBUG MODE: Testing Info level logging");
        logger.Error("DEBUG MODE: Testing Error level logging");
        logger.Debug("DEBUG MODE: Testing Debug level logging");
        
        Console.WriteLine("\n=== CHECKING LOG DIRECTORIES ===");
        var logPaths = new[]
        {
            "logs",
            "logs/archive"
        };
        
        foreach (var path in logPaths)
        {
            var fullPath = Path.GetFullPath(path);
            var exists = Directory.Exists(fullPath);
            Console.WriteLine($"Directory {path}: {(exists ? "EXISTS" : "MISSING")} ({fullPath})");
            
            if (exists)
            {
                var files = Directory.GetFiles(fullPath, "*", SearchOption.TopDirectoryOnly);
                Console.WriteLine($"  Files in directory: {files.Length}");
                foreach (var file in files.Take(5))
                {
                    var fileInfo = new FileInfo(file);
                    Console.WriteLine($"    {Path.GetFileName(file)} ({fileInfo.Length} bytes, {fileInfo.LastWriteTime})");
                }
            }
        }
        
        Console.WriteLine("\n=== TESTING REGISTRY SCANNER ===");
        try
        {
            using var scanner = new AutodeskRegistryScanner();
            var entries = await scanner.ScanRegistryAsync();
            Console.WriteLine($"Found {entries.Count} registry entries");
            
            if (entries.Count > 0)
            {
                Console.WriteLine("\n=== TESTING VALIDATION ===");
                var isValid = scanner.ValidateEntries(entries);
                Console.WriteLine($"Validation result: {isValid}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Registry scanner error: {ex.Message}");
        }
        
        Console.WriteLine("\n=== DEBUG MODE COMPLETE ===");
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
        
        return 0;
    }
#pragma warning restore Spectre1000
    
    /// <summary>
    /// Formats a file size into a human-readable string.
    /// </summary>
    /// <param name="bytes">The size in bytes.</param>
    /// <returns>Formatted file size string.</returns>
    private static string FormatFileSize(long bytes)
    {
        const int scale = 1024;
        string[] orders = ["GB", "MB", "KB", "Bytes"];
        long max = (long)Math.Pow(scale, orders.Length - 1);

        foreach (var order in orders)
        {
            if (bytes > max)
            {
                return $"{decimal.Divide(bytes, max):##.##} {order}";
            }

            max /= scale;
        }
        
        return "0 Bytes";
    }
}
