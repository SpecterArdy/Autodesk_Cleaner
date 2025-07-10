using Autodesk_Cleaner.Core;
using System.Security.Principal;
using NLog;

namespace Autodesk_Cleaner;

/// <summary>
/// Main entry point for the Autodesk Registry Cleaner console application.
/// </summary>
internal static class Program
{
    private static readonly ConsoleColor OriginalForegroundColor = Console.ForegroundColor;
    private static readonly ConsoleColor OriginalBackgroundColor = Console.BackgroundColor;
    private static Logger? _logger;

    /// <summary>
    /// Main entry point for the application.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <returns>Exit code (0 for success, non-zero for error).</returns>
    private static async Task<int> Main(string[] args)
    {
        try
        {
            // Initialize logging first
            LoggingConfiguration.InitializeLogging();
            _logger = LogManager.GetCurrentClassLogger();
            
            Console.Title = "Autodesk Registry Cleaner v1.0.0";
            
            _logger.Info("Application started with arguments: {@Args}", args);
            
            WriteHeader();
            
            // Check for administrator privileges
            if (!IsRunningAsAdministrator())
            {
                _logger.Error("Application not running as Administrator");
                WriteError("This application must be run as Administrator to modify registry and system files.");
                WriteWarning("Please restart the application with administrator privileges.");
                WriteInfo("Press any key to exit...");
                Console.ReadKey();
                return 1;
            }
            
            _logger.Info("Administrator privileges confirmed");

            // Parse command line arguments
            var config = ParseArguments(args);
            _logger.Info("Configuration parsed: {@Config}", config);
            
            // Display configuration
            DisplayConfiguration(config);
            
            // Get user confirmation
            if (!config.DryRun && !GetUserConfirmation())
            {
                _logger.Info("Operation cancelled by user");
                WriteInfo("Operation cancelled by user.");
                return 0;
            }

            // Execute the cleaning operation
            var exitCode = await ExecuteCleaningOperationAsync(config);
            
            WriteInfo("Press any key to exit...");
            Console.ReadKey();
            
            return exitCode;
        }
        catch (Exception ex)
        {
            _logger?.Fatal(ex, "Critical error in main application");
            WriteError($"Critical error: {ex.Message}");
            WriteError(ex.StackTrace ?? "No stack trace available");
            
            WriteInfo("Press any key to exit...");
            Console.ReadKey();
            
            return -1;
        }
        finally
        {
            _logger?.Info("Application shutdown");
            LogManager.Shutdown();
            
            // Restore original console colors
            Console.ForegroundColor = OriginalForegroundColor;
            Console.BackgroundColor = OriginalBackgroundColor;
        }
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
            WriteInfo("Starting Autodesk cleanup operation...");
            Console.WriteLine();

            // Phase 1: Registry Cleanup
            WriteStep("Phase 1: Registry Cleanup");
            using var registryScanner = new AutodeskRegistryScanner(config);
            
            WriteInfo("Scanning registry for Autodesk entries...");
            var registryEntries = await registryScanner.ScanRegistryAsync();
            WriteSuccess($"Found {registryEntries.Count} registry entries to clean.");

            if (registryEntries.Count > 0)
            {
                WriteInfo("Registry entries to be removed:");
                foreach (var entry in registryEntries.Take(10)) // Show first 10
                {
                    WriteDetail($"  • {entry.DisplayName}");
                }
                
                if (registryEntries.Count > 10)
                {
                    WriteDetail($"  ... and {registryEntries.Count - 10} more entries");
                }

                var registryResult = await registryScanner.RemoveEntriesAsync(registryEntries);
                DisplayRemovalResult("Registry", registryResult);
                
                if (!registryResult.IsSuccessful)
                {
                    overallSuccess = false;
                    totalErrors.AddRange(registryResult.Errors);
                }
            }
            else
            {
                WriteInfo("No registry entries found to clean.");
            }

            Console.WriteLine();

            // Phase 2: File System Cleanup
            WriteStep("Phase 2: File System Cleanup");
            using var fileSystemCleaner = new AutodeskFileSystemCleaner(config);
            
            WriteInfo("Scanning file system for Autodesk files and directories...");
            var fileSystemEntries = await fileSystemCleaner.ScanFileSystemAsync();
            WriteSuccess($"Found {fileSystemEntries.Count} file system entries to clean.");

            if (fileSystemEntries.Count > 0)
            {
                WriteInfo("File system entries to be removed:");
                foreach (var entry in fileSystemEntries.Take(10)) // Show first 10
                {
                    var sizeText = entry.EntryType == FileSystemEntryType.File 
                        ? $" ({FormatFileSize(entry.Size)})" 
                        : " (Directory)";
                    WriteDetail($"  • {entry.Path}{sizeText}");
                }
                
                if (fileSystemEntries.Count > 10)
                {
                    WriteDetail($"  ... and {fileSystemEntries.Count - 10} more entries");
                }

                var fileSystemResult = await fileSystemCleaner.RemoveEntriesAsync(fileSystemEntries);
                DisplayRemovalResult("File System", fileSystemResult);
                
                if (!fileSystemResult.IsSuccessful)
                {
                    overallSuccess = false;
                    totalErrors.AddRange(fileSystemResult.Errors);
                }
            }
            else
            {
                WriteInfo("No file system entries found to clean.");
            }

            Console.WriteLine();

            // Summary
            WriteSummary(overallSuccess, totalErrors, config.DryRun);

            return overallSuccess ? 0 : 1;
        }
        catch (Exception ex)
        {
            WriteError($"Operation failed: {ex.Message}");
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
        if (result.IsSuccessful)
        {
            WriteSuccess($"{phaseName} cleanup completed successfully!");
        }
        else
        {
            WriteWarning($"{phaseName} cleanup completed with some errors.");
        }

        WriteDetail($"  • Total entries: {result.TotalEntries}");
        WriteDetail($"  • Successfully removed: {result.SuccessfulRemovals}");
        WriteDetail($"  • Failed removals: {result.FailedRemovals}");
        WriteDetail($"  • Success rate: {result.SuccessRate:F1}%");
        WriteDetail($"  • Duration: {result.Duration.TotalSeconds:F2} seconds");

        if (result.Errors.Count > 0)
        {
            WriteWarning($"  • Errors encountered: {result.Errors.Count}");
            foreach (var error in result.Errors.Take(5)) // Show first 5 errors
            {
                WriteError($"    - {error}");
            }
            
            if (result.Errors.Count > 5)
            {
                WriteDetail($"    ... and {result.Errors.Count - 5} more errors");
            }
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
        WriteStep("Operation Summary");
        
        if (isDryRun)
        {
            WriteInfo("DRY RUN COMPLETED - No actual changes were made to your system.");
        }
        else if (overallSuccess)
        {
            WriteSuccess("Autodesk cleanup completed successfully!");
            WriteInfo("All Autodesk registry entries and files have been removed from your system.");
            WriteInfo("You can now proceed with a fresh installation of Autodesk products.");
        }
        else
        {
            WriteWarning("Autodesk cleanup completed with some errors.");
            WriteInfo($"Total errors encountered: {totalErrors.Count}");
            WriteInfo("Some Autodesk entries may still remain on your system.");
            WriteInfo("You may need to manually remove remaining entries or run the tool again.");
        }

        Console.WriteLine();
        WriteInfo("IMPORTANT NOTES:");
        WriteDetail("• Restart your computer before installing new Autodesk products");
        WriteDetail("• Clear your browser cache if using web-based installers");
        WriteDetail("• Temporarily disable antivirus during Autodesk installation");
        WriteDetail("• Run Windows Update before installing new Autodesk products");
    }

    /// <summary>
    /// Parses command line arguments into a ScannerConfig.
    /// </summary>
    /// <param name="args">Command line arguments.</param>
    /// <returns>Parsed configuration.</returns>
    private static ScannerConfig ParseArguments(string[] args)
    {
        var dryRun = args.Contains("--dry-run") || args.Contains("-d");
        var noBackup = args.Contains("--no-backup") || args.Contains("-n");
        var userOnly = args.Contains("--user-only") || args.Contains("-u");
        var systemOnly = args.Contains("--system-only") || args.Contains("-s");

        var backupPath = ScannerConfig.Default.BackupPath;
        var backupIndex = Array.IndexOf(args, "--backup-path");
        if (backupIndex >= 0 && backupIndex + 1 < args.Length)
        {
            backupPath = args[backupIndex + 1];
        }

        return new ScannerConfig(
            CreateBackup: !noBackup,
            BackupPath: backupPath,
            DryRun: dryRun,
            IncludeUserHive: !systemOnly,
            IncludeLocalMachine: !userOnly,
            MaxDepth: 10);
    }

    /// <summary>
    /// Displays the current configuration.
    /// </summary>
    /// <param name="config">The configuration to display.</param>
    private static void DisplayConfiguration(ScannerConfig config)
    {
        WriteStep("Configuration");
        WriteDetail($"Dry Run: {(config.DryRun ? "Yes" : "No")}");
        WriteDetail($"Create Backup: {(config.CreateBackup ? "Yes" : "No")}");
        if (config.CreateBackup)
        {
            WriteDetail($"Backup Path: {config.BackupPath}");
        }
        WriteDetail($"Include User Registry: {(config.IncludeUserHive ? "Yes" : "No")}");
        WriteDetail($"Include System Registry: {(config.IncludeLocalMachine ? "Yes" : "No")}");
        WriteDetail($"Max Scan Depth: {config.MaxDepth}");
        Console.WriteLine();
    }

    /// <summary>
    /// Gets user confirmation before proceeding with the operation.
    /// </summary>
    /// <returns>True if the user confirmed, false otherwise.</returns>
    private static bool GetUserConfirmation()
    {
        WriteWarning("WARNING: This operation will permanently remove ALL Autodesk products and data from your system!");
        WriteWarning("Make sure you have backed up any important project files before proceeding.");
        Console.WriteLine();
        
        WriteInfo("Do you want to proceed with the cleanup? (y/N): ");
        var response = Console.ReadLine()?.Trim().ToLowerInvariant();
        return response == "y" || response == "yes";
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
    /// Displays the application header.
    /// </summary>
    private static void WriteHeader()
    {
        Console.Clear();
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("╔══════════════════════════════════════════════════════════════════════════════╗");
        Console.WriteLine("║                        AUTODESK REGISTRY CLEANER v1.0.0                    ║");
        Console.WriteLine("║                     Modular Tool for Complete Autodesk Removal              ║");
        Console.WriteLine("╚══════════════════════════════════════════════════════════════════════════════╝");
        Console.ForegroundColor = OriginalForegroundColor;
        Console.WriteLine();
        
        WriteInfo("This tool will completely remove all Autodesk products from your system including:");
        WriteDetail("• Registry entries (HKLM and HKCU)");
        WriteDetail("• Program files and directories");
        WriteDetail("• User profile data");
        WriteDetail("• Temporary files");
        WriteDetail("• License files");
        Console.WriteLine();
        
        WriteInfo("Command line options:");
        WriteDetail("  --dry-run, -d       : Preview changes without making them");
        WriteDetail("  --no-backup, -n     : Skip creating backups");
        WriteDetail("  --user-only, -u     : Clean only user registry and files");
        WriteDetail("  --system-only, -s   : Clean only system registry and files");
        WriteDetail("  --backup-path PATH  : Specify custom backup location");
        Console.WriteLine();
    }

    /// <summary>
    /// Writes a step header to the console.
    /// </summary>
    /// <param name="message">The step message.</param>
    private static void WriteStep(string message)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"=== {message} ===");
        Console.ForegroundColor = OriginalForegroundColor;
    }

    /// <summary>
    /// Writes an informational message to the console.
    /// </summary>
    /// <param name="message">The message to write.</param>
    private static void WriteInfo(string message)
    {
        Console.ForegroundColor = ConsoleColor.White;
        Console.Write(message);
        Console.ForegroundColor = OriginalForegroundColor;
    }

    /// <summary>
    /// Writes a success message to the console.
    /// </summary>
    /// <param name="message">The message to write.</param>
    private static void WriteSuccess(string message)
    {
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"✓ {message}");
        Console.ForegroundColor = OriginalForegroundColor;
    }

    /// <summary>
    /// Writes a warning message to the console.
    /// </summary>
    /// <param name="message">The message to write.</param>
    private static void WriteWarning(string message)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"⚠ {message}");
        Console.ForegroundColor = OriginalForegroundColor;
    }

    /// <summary>
    /// Writes an error message to the console.
    /// </summary>
    /// <param name="message">The message to write.</param>
    private static void WriteError(string message)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"✗ {message}");
        Console.ForegroundColor = OriginalForegroundColor;
    }

    /// <summary>
    /// Writes a detail message to the console.
    /// </summary>
    /// <param name="message">The message to write.</param>
    private static void WriteDetail(string message)
    {
        Console.ForegroundColor = ConsoleColor.Gray;
        Console.WriteLine(message);
        Console.ForegroundColor = OriginalForegroundColor;
    }

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
