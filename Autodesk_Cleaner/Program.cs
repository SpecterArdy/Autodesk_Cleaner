using Autodesk_Cleaner.Core;
using Microsoft.Win32;
using System.Diagnostics;
using System.Security.AccessControl;
using System.Security.Principal;
using System.ServiceProcess;
using System.Text.RegularExpressions;
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
    private const string AdskNetworkLicenseServiceName = "AdskNLM";
    private const string ThreeDsMax2027ProductKey = "128S1";
    private const string ThreeDsMax2027ProductVersion = "2027.0.0.F";
    private static readonly Regex AutodeskServicePattern = new(
        @"(autodesk|adsk|3ds\s*max|maya|revit|inventor|navisworks|motionbuilder|mudbox|alias|genuine)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly HashSet<string> KnownAutodeskServiceNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "AdskLicensingService",
        "AdskAccessServiceHost",
        "AdAppMgrSvc",
        "Mi-Service",
        "AdskNLM",
        "AGSService",
        "AutodeskDesktopAppService"
    };

    private readonly record struct LicenseRepairStep(string Step, string Status, string Details);
    private readonly record struct AutodeskServiceInfo(string ServiceName, string DisplayName, ServiceControllerStatus Status);
    private readonly record struct ServiceCleanupResult(int TotalServices, int RemovedServices, IReadOnlyCollection<string> Errors);

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
                WaitForUserKeyIfInteractive();
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
            WaitForUserKeyIfInteractive();
            
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
                case InteractiveMenu.MenuOption.Repair3dsMax2027Licensing:
                    await HandleRepair3dsMax2027LicensingAsync();
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
    /// Repairs stale Autodesk network-license settings that can block 3ds Max 2027 named-user sign-in.
    /// </summary>
    private static async Task HandleRepair3dsMax2027LicensingAsync()
    {
        var repairSteps = await AnsiConsole.Status()
            .StartAsync("[yellow]Repairing stale Autodesk licensing configuration...[/]", async ctx =>
            {
                ctx.Status("Running Autodesk licensing helper...");
                var steps = new List<LicenseRepairStep>
                {
                    await Reset3dsMax2027NamedUserLicensingAsync()
                };

                ctx.Status("Clearing Autodesk licensing service cache...");
                steps.Add(RemoveAdskLicensingServiceCache());

                ctx.Status("Clearing FLEXlm network-license overrides...");
                steps.Add(RemoveRegistryValueIfPresent(
                    RegistryHive.CurrentUser,
                    RegistryView.Default,
                    @"Software\FLEXlm License Manager",
                    "ADSKFLEX_LICENSE_FILE",
                    @"HKCU\Software\FLEXlm License Manager\ADSKFLEX_LICENSE_FILE"));
                steps.Add(RemoveRegistryValueIfPresent(
                    RegistryHive.LocalMachine,
                    RegistryView.Registry64,
                    @"Software\FLEXlm License Manager",
                    "ADSKFLEX_LICENSE_FILE",
                    @"HKLM\Software\FLEXlm License Manager\ADSKFLEX_LICENSE_FILE"));
                steps.Add(RemoveRegistryValueIfPresent(
                    RegistryHive.LocalMachine,
                    RegistryView.Registry32,
                    @"Software\FLEXlm License Manager",
                    "ADSKFLEX_LICENSE_FILE",
                    @"HKLM\Software\Wow6432Node\FLEXlm License Manager\ADSKFLEX_LICENSE_FILE"));

                ctx.Status("Updating Autodesk Network License Manager service...");
                steps.Add(await SetServiceStartupManualAsync(AdskNetworkLicenseServiceName));

                return steps;
            });

        DisplayLicenseRepairSummary(repairSteps);
    }

    /// <summary>
    /// Runs Autodesk's licensing helper to switch 3ds Max 2027 back to named-user mode.
    /// </summary>
    private static async Task<LicenseRepairStep> Reset3dsMax2027NamedUserLicensingAsync()
    {
        var helperPath = GetAdskLicensingHelperPath();
        if (!File.Exists(helperPath))
        {
            return new LicenseRepairStep(
                "Reset named-user license mode",
                "Skipped",
                $"Licensing helper not found at {helperPath}");
        }

        var arguments = $"change -pk {ThreeDsMax2027ProductKey} -pv {ThreeDsMax2027ProductVersion} -lm \"\"";
        var (success, standardOutput, standardError) = await RunProcessAsync(helperPath, arguments);
        var details = string.IsNullOrWhiteSpace(standardOutput)
            ? "Autodesk licensing helper completed."
            : standardOutput.Trim();

        if (!success)
        {
            details = string.IsNullOrWhiteSpace(standardError)
                ? "Autodesk licensing helper returned a non-zero exit code."
                : standardError.Trim();
        }

        return new LicenseRepairStep(
            "Reset named-user license mode",
            success ? "Applied" : "Failed",
            details);
    }

    /// <summary>
    /// Removes a registry value if it exists.
    /// </summary>
    private static LicenseRepairStep RemoveRegistryValueIfPresent(
        RegistryHive hive,
        RegistryView view,
        string subKeyPath,
        string valueName,
        string displayPath)
    {
        try
        {
            using var baseKey = RegistryKey.OpenBaseKey(hive, view);
            using var subKey = baseKey.OpenSubKey(subKeyPath, writable: true);
            if (subKey is null)
            {
                return new LicenseRepairStep("Clear stale FLEXlm override", "Not present", displayPath);
            }

            var existingValue = subKey.GetValue(valueName);
            if (existingValue is null)
            {
                return new LicenseRepairStep("Clear stale FLEXlm override", "Not present", displayPath);
            }

            var existingValueText = Convert.ToString(existingValue) ?? "<non-string>";
            subKey.DeleteValue(valueName, throwOnMissingValue: false);

            return new LicenseRepairStep(
                "Clear stale FLEXlm override",
                "Removed",
                $"{displayPath} was set to '{existingValueText}'");
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "Failed to remove registry value {DisplayPath}", displayPath);
            return new LicenseRepairStep(
                "Clear stale FLEXlm override",
                "Failed",
                $"{displayPath}: {ex.Message}");
        }
    }

    /// <summary>
    /// Removes the 3ds Max 2027 licensing service cache used by Autodesk's local licensing service.
    /// </summary>
    private static LicenseRepairStep RemoveAdskLicensingServiceCache()
    {
        const string licensingServiceRoot = @"C:\ProgramData\Autodesk\AdskLicensingService";

        try
        {
            if (!Directory.Exists(licensingServiceRoot))
            {
                return new LicenseRepairStep(
                    "Clear AdskLicensingService product cache",
                    "Not present",
                    licensingServiceRoot);
            }

            var matchingDirectories = Directory
                .GetDirectories(licensingServiceRoot, $"{ThreeDsMax2027ProductKey}*", SearchOption.TopDirectoryOnly)
                .ToList();

            if (matchingDirectories.Count == 0)
            {
                return new LicenseRepairStep(
                    "Clear AdskLicensingService product cache",
                    "Not present",
                    $"No {ThreeDsMax2027ProductKey} cache folders found under {licensingServiceRoot}");
            }

            foreach (var directoryPath in matchingDirectories)
            {
                Directory.Delete(directoryPath, recursive: true);
            }

            return new LicenseRepairStep(
                "Clear AdskLicensingService product cache",
                "Removed",
                string.Join(", ", matchingDirectories.Select(Path.GetFileName)));
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "Failed to remove Autodesk licensing service cache for product key {ProductKey}", ThreeDsMax2027ProductKey);
            return new LicenseRepairStep(
                "Clear AdskLicensingService product cache",
                "Failed",
                ex.Message);
        }
    }

    /// <summary>
    /// Sets a Windows service to Manual startup if it is installed.
    /// </summary>
    private static async Task<LicenseRepairStep> SetServiceStartupManualAsync(string serviceName)
    {
        try
        {
            var serviceExists = ServiceController.GetServices()
                .Any(service => string.Equals(service.ServiceName, serviceName, StringComparison.OrdinalIgnoreCase));

            if (!serviceExists)
            {
                return new LicenseRepairStep(
                    "Set AdskNLM startup to Manual",
                    "Not installed",
                    $"{serviceName} is not registered on this machine.");
            }

            var (success, standardOutput, standardError) =
                await RunProcessAsync("sc.exe", $"config \"{serviceName}\" start= demand");

            var details = success
                ? $"Updated {serviceName} startup type to Manual."
                : string.IsNullOrWhiteSpace(standardError) ? standardOutput.Trim() : standardError.Trim();

            return new LicenseRepairStep(
                "Set AdskNLM startup to Manual",
                success ? "Applied" : "Failed",
                details);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "Failed to update startup type for service {ServiceName}", serviceName);
            return new LicenseRepairStep(
                "Set AdskNLM startup to Manual",
                "Failed",
                ex.Message);
        }
    }

    /// <summary>
    /// Gets the Autodesk licensing helper path used to reset product license mode.
    /// </summary>
    private static string GetAdskLicensingHelperPath()
    {
        var commonProgramFilesX86 = Environment.GetEnvironmentVariable("CommonProgramFiles(x86)");
        if (string.IsNullOrWhiteSpace(commonProgramFilesX86))
        {
            commonProgramFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86);
        }

        return Path.Combine(
            commonProgramFilesX86,
            "Autodesk Shared",
            "AdskLicensing",
            "Current",
            "helper",
            "AdskLicensingInstHelper.exe");
    }

    /// <summary>
    /// Runs a process and captures its standard output and error.
    /// </summary>
    private static async Task<(bool Success, string StandardOutput, string StandardError)> RunProcessAsync(
        string fileName,
        string arguments)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        using var process = Process.Start(startInfo);
        if (process is null)
        {
            return (false, string.Empty, $"Failed to start process: {fileName}");
        }

        var standardOutputTask = process.StandardOutput.ReadToEndAsync();
        var standardErrorTask = process.StandardError.ReadToEndAsync();

        await process.WaitForExitAsync();

        return (process.ExitCode == 0, await standardOutputTask, await standardErrorTask);
    }

    /// <summary>
    /// Displays the results of the licensing repair workflow.
    /// </summary>
    private static void DisplayLicenseRepairSummary(IReadOnlyCollection<LicenseRepairStep> repairSteps)
    {
        var table = new Table()
        {
            Border = TableBorder.Rounded,
            BorderStyle = new Style(Color.Yellow)
        };

        table.AddColumn("[bold]Step[/]");
        table.AddColumn("[bold]Status[/]");
        table.AddColumn("[bold]Details[/]");

        foreach (var repairStep in repairSteps)
        {
            var statusMarkup = repairStep.Status switch
            {
                "Applied" or "Removed" => $"[green]{repairStep.Status}[/]",
                "Skipped" or "Not present" or "Not installed" => $"[yellow]{repairStep.Status}[/]",
                "Failed" => $"[red]{repairStep.Status}[/]",
                _ => repairStep.Status
            };

            table.AddRow(repairStep.Step, statusMarkup, Markup.Escape(repairStep.Details));
        }

        var panel = new Panel(table)
        {
            Header = new PanelHeader(" [bold yellow]3DS MAX 2027 LICENSING REPAIR[/] "),
            Border = BoxBorder.Double,
            BorderStyle = new Style(Color.Yellow)
        };

        AnsiConsole.Write(panel);

        var hasFailure = repairSteps.Any(step => string.Equals(step.Status, "Failed", StringComparison.OrdinalIgnoreCase));
        var nextStepsPanel = new Panel(new Markup(
            "[bold cyan]Next steps[/]\n\n" +
            "1. Reboot Windows.\n" +
            "2. Launch 3ds Max 2027.\n" +
            "3. Use the normal Autodesk sign-in / named-user flow.\n\n" +
            (hasFailure
                ? "[yellow]One or more repair steps failed. Review the details above before testing 3ds Max again.[/]"
                : "[green]The stale FLEXlm override and local network-license settings have been repaired.[/]")))
        {
            Header = new PanelHeader(" [bold cyan]FOLLOW-UP[/] "),
            Border = BoxBorder.Rounded,
            BorderStyle = new Style(Color.Blue)
        };

        AnsiConsole.Write(nextStepsPanel);
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

            // Phase 0: Service Cleanup
            AnsiConsole.Write(new Rule("[bold blue]Phase 0: Service Cleanup[/]"));

            var serviceCleanupResult = await CleanupAutodeskServicesAsync(config.DryRun);
            DisplayServiceCleanupResult(serviceCleanupResult, config.DryRun);

            if (!config.DryRun && serviceCleanupResult.Errors.Count > 0)
            {
                overallSuccess = false;
                totalErrors.AddRange(serviceCleanupResult.Errors);
            }

            // Phase 1: Licensing Cleanup
            AnsiConsole.Write(new Rule("[bold blue]Phase 1: Licensing Cleanup[/]"));

            AnsiConsole.MarkupLine("[green]Cleaning Autodesk licensing overrides...[/]");
            AutodeskLicensingCleanupResult licensingResult;
            var licensingCleaner = new AutodeskLicensingCleaner(config);
            licensingResult = await licensingCleaner.CleanupAsync();

            DisplayLicensingCleanupResult(licensingResult, config.DryRun);

            if (!licensingResult.IsSuccessful)
            {
                overallSuccess = false;
                totalErrors.AddRange(licensingResult.Errors);
            }

            // Phase 2: MSI Registration Cleanup
            AnsiConsole.Write(new Rule("[bold blue]Phase 2: Installer Registration Cleanup[/]"));

            AnsiConsole.MarkupLine("[green]Cleaning Autodesk Windows Installer registrations...[/]");
            BrokenMsiCleanupResult msiCleanupResult;
            var msiRegistrationCleaner = new BrokenMsiRegistrationCleaner(config);
            msiCleanupResult = await msiRegistrationCleaner.CleanupAsync();

            DisplayMsiCleanupResult(msiCleanupResult, config.DryRun);

            if (!msiCleanupResult.IsSuccessful)
            {
                overallSuccess = false;
                totalErrors.AddRange(msiCleanupResult.Errors);
            }

            // Phase 3: Registry Cleanup
            AnsiConsole.Write(new Rule("[bold blue]Phase 3: Registry Cleanup[/]"));
            
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
                
                AnsiConsole.MarkupLine("[yellow]Processing registry entries...[/]");
                RemovalResult registryResult;
                using (var registryScanner = new AutodeskRegistryScanner(config))
                {
                    registryResult = await registryScanner.RemoveEntriesAsync(registryEntries);
                }
                    
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

            // Phase 4: File System Cleanup
            AnsiConsole.Write(new Rule("[bold blue]Phase 4: File System Cleanup[/]"));
            
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
                
                AnsiConsole.MarkupLine("[yellow]Processing file system entries...[/]");
                RemovalResult fileSystemResult;
                using (var fileSystemCleaner = new AutodeskFileSystemCleaner(config))
                {
                    fileSystemResult = await fileSystemCleaner.RemoveEntriesAsync(fileSystemEntries);
                }
                    
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
    /// Finds and removes installed Autodesk-related Windows services.
    /// </summary>
    private static async Task<ServiceCleanupResult> CleanupAutodeskServicesAsync(bool dryRun)
    {
        var services = DiscoverAutodeskServices();
        var errors = new List<string>();

        if (services.Count == 0)
        {
            AnsiConsole.MarkupLine("[dim]No Autodesk services are currently registered.[/]");
            return new ServiceCleanupResult(0, 0, errors);
        }

        AnsiConsole.MarkupLine($"[green]Found {services.Count} Autodesk-related services.[/]");

        var removedServices = 0;
        var index = 0;
        foreach (var service in services)
        {
            index++;

            if (dryRun)
            {
                AnsiConsole.MarkupLine(
                    $"[dim][DRY RUN] Would remove service {index}/{services.Count}:[/] {Markup.Escape(service.ServiceName)} [dim]({Markup.Escape(service.DisplayName)})[/]");
                removedServices++;
                continue;
            }

            AnsiConsole.MarkupLine(
                $"[yellow]Removing service {index}/{services.Count}:[/] {Markup.Escape(service.ServiceName)} [dim]({Markup.Escape(service.DisplayName)})[/]");

            var removed = await RemoveServiceAsync(service);
            if (removed)
            {
                removedServices++;
                AnsiConsole.MarkupLine($"[green]Removed service:[/] {Markup.Escape(service.ServiceName)}");
            }
            else
            {
                var error = $"Failed to remove service: {service.ServiceName}";
                errors.Add(error);
                AnsiConsole.MarkupLine($"[red]{Markup.Escape(error)}[/]");
            }
        }

        return new ServiceCleanupResult(services.Count, removedServices, errors);
    }

    /// <summary>
    /// Displays service cleanup results.
    /// </summary>
    private static void DisplayServiceCleanupResult(ServiceCleanupResult result, bool isDryRun)
    {
        var table = new Table()
        {
            Border = TableBorder.Rounded,
            BorderStyle = result.Errors.Count == 0 ? new Style(Color.Green) : new Style(Color.Yellow)
        };

        table.AddColumn("[bold]Metric[/]");
        table.AddColumn("[bold]Value[/]");
        table.AddRow("Services Found", result.TotalServices.ToString());
        table.AddRow(isDryRun ? "Services Planned" : "Services Removed", $"[green]{result.RemovedServices}[/]");
        table.AddRow("Failed Removals", result.Errors.Count > 0 ? $"[red]{result.Errors.Count}[/]" : "0");

        var panel = new Panel(table)
        {
            Header = new PanelHeader($" [bold {(result.Errors.Count == 0 ? "green" : "yellow")}]SERVICE CLEANUP RESULTS[/] "),
            Border = BoxBorder.Rounded
        };

        AnsiConsole.Write(panel);

        if (result.Errors.Count > 0)
        {
            var errorPanel = new Panel(string.Join("\n", result.Errors.Select(error => $"[red]• {Markup.Escape(error)}[/]")))
            {
                Header = new PanelHeader(" [bold red]Service Cleanup Errors[/] "),
                Border = BoxBorder.Rounded,
                BorderStyle = new Style(Color.Red)
            };

            AnsiConsole.Write(errorPanel);
        }
    }

    /// <summary>
    /// Displays Autodesk licensing cleanup results.
    /// </summary>
    private static void DisplayLicensingCleanupResult(AutodeskLicensingCleanupResult result, bool isDryRun)
    {
        var table = new Table()
        {
            Border = TableBorder.Rounded,
            BorderStyle = result.IsSuccessful ? new Style(Color.Green) : new Style(Color.Yellow)
        };

        table.AddColumn("[bold]Step[/]");
        table.AddColumn("[bold]Status[/]");
        table.AddColumn("[bold]Details[/]");

        foreach (var step in result.Steps)
        {
            var statusColor = step.Status switch
            {
                "Removed" or "Applied" => "green",
                "Planned" => "yellow",
                "Not present" or "Not installed" or "Skipped" => "grey",
                _ => "red"
            };

            table.AddRow(
                Markup.Escape(step.Step),
                $"[{statusColor}]{Markup.Escape(step.Status)}[/]",
                Markup.Escape(step.Details));
        }

        var panel = new Panel(table)
        {
            Header = new PanelHeader($" [bold {(result.IsSuccessful ? "green" : "yellow")}]LICENSING CLEANUP {(isDryRun ? "PLAN" : "RESULTS")}[/] "),
            Border = BoxBorder.Rounded
        };

        AnsiConsole.Write(panel);
    }

    /// <summary>
    /// Displays Autodesk MSI cleanup results.
    /// </summary>
    private static void DisplayMsiCleanupResult(BrokenMsiCleanupResult result, bool isDryRun)
    {
        var summaryTable = new Table()
        {
            Border = TableBorder.Rounded,
            BorderStyle = result.IsSuccessful ? new Style(Color.Green) : new Style(Color.Yellow)
        };

        summaryTable.AddColumn("[bold]Metric[/]");
        summaryTable.AddColumn("[bold]Value[/]");
        summaryTable.AddRow("Autodesk MSI Registrations", result.Candidates.Count.ToString());
        summaryTable.AddRow(isDryRun ? "Planned Removals" : "Successful Removals", result.SuccessfulRemovals.ToString());
        summaryTable.AddRow("Running msiexec", result.RunningMsiexec.Count.ToString());
        summaryTable.AddRow("Errors", result.Errors.Count.ToString());

        var summaryPanel = new Panel(summaryTable)
        {
            Header = new PanelHeader($" [bold {(result.IsSuccessful ? "green" : "yellow")}]INSTALLER CLEANUP {(isDryRun ? "PLAN" : "RESULTS")}[/] "),
            Border = BoxBorder.Rounded
        };

        AnsiConsole.Write(summaryPanel);

        if (result.Candidates.Count > 0)
        {
            var candidateTable = new Table()
            {
                Border = TableBorder.Rounded,
                BorderStyle = new Style(Color.Blue)
            };

            candidateTable.AddColumn("[bold]Product[/]");
            candidateTable.AddColumn("[bold]Packed Key[/]");
            candidateTable.AddColumn("[bold]Reason[/]");

            foreach (var candidate in result.Candidates.Take(12))
            {
                candidateTable.AddRow(
                    Markup.Escape(candidate.DisplayName),
                    Markup.Escape(candidate.PackedProductCode),
                    Markup.Escape(candidate.Reasons.FirstOrDefault() ?? "Autodesk installer registration"));
            }

            if (result.Candidates.Count > 12)
            {
                candidateTable.AddRow($"[dim]... and {result.Candidates.Count - 12} more[/]", string.Empty, string.Empty);
            }

            AnsiConsole.Write(new Panel(candidateTable)
            {
                Header = new PanelHeader(" [bold blue]Installer Registrations Selected[/] "),
                Border = BoxBorder.Rounded
            });
        }

        if (result.RunningMsiexec.Count > 0)
        {
            var msiexecLines = result.RunningMsiexec.Select(process =>
            {
                var tag = process.IsAutodeskRelated ? "[yellow]Autodesk-related[/]" : "[red]Non-Autodesk[/]";
                var commandLine = string.IsNullOrWhiteSpace(process.CommandLine) ? "<no command line>" : process.CommandLine;
                return $"{tag} PID {process.ProcessId}: {Markup.Escape(commandLine)}";
            });

            AnsiConsole.Write(new Panel(string.Join('\n', msiexecLines))
            {
                Header = new PanelHeader(" [bold yellow]Running msiexec Processes[/] "),
                Border = BoxBorder.Rounded,
                BorderStyle = new Style(Color.Yellow)
            });
        }

        if (result.Errors.Count > 0)
        {
            AnsiConsole.Write(new Panel(string.Join('\n', result.Errors.Select(error => $"[red]• {Markup.Escape(error)}[/]")))
            {
                Header = new PanelHeader(" [bold red]Installer Cleanup Errors[/] "),
                Border = BoxBorder.Rounded,
                BorderStyle = new Style(Color.Red)
            });
        }
    }

    /// <summary>
    /// Discovers installed Autodesk-related Windows services.
    /// </summary>
    private static List<AutodeskServiceInfo> DiscoverAutodeskServices()
    {
        var services = new List<AutodeskServiceInfo>();

        foreach (var service in ServiceController.GetServices())
        {
            using (service)
            {
                var serviceName = service.ServiceName ?? string.Empty;
                var displayName = service.DisplayName ?? string.Empty;

                if (KnownAutodeskServiceNames.Contains(serviceName) ||
                    AutodeskServicePattern.IsMatch(serviceName) ||
                    AutodeskServicePattern.IsMatch(displayName))
                {
                    services.Add(new AutodeskServiceInfo(serviceName, displayName, service.Status));
                }
            }
        }

        return services
            .OrderBy(service => service.ServiceName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>
    /// Removes a Windows service from the SCM and service registry hive.
    /// </summary>
    private static async Task<bool> RemoveServiceAsync(AutodeskServiceInfo service)
    {
        try
        {
            await StopServiceForRemovalAsync(service.ServiceName);
            await RunProcessAsync("sc.exe", $"delete \"{service.ServiceName}\"");

            var registryDeleted = TryDeleteServiceRegistryKey(service.ServiceName);
            var stillRegistered = IsServiceRegistered(service.ServiceName);

            if (stillRegistered)
            {
                _logger?.Warn("Service {ServiceName} still appears registered after deletion attempt", service.ServiceName);
            }

            return !stillRegistered || registryDeleted;
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "Failed to remove service {ServiceName}", service.ServiceName);
            return false;
        }
    }

    /// <summary>
    /// Stops a service and force-kills its backing process if needed.
    /// </summary>
    private static async Task StopServiceForRemovalAsync(string serviceName)
    {
        try
        {
            using var service = new ServiceController(serviceName);
            if (service.Status is ServiceControllerStatus.Stopped or ServiceControllerStatus.StopPending)
            {
                await WaitForServiceStoppedAsync(serviceName, TimeSpan.FromSeconds(15));
                return;
            }

            if (service.CanStop)
            {
                service.Stop();
                await WaitForServiceStoppedAsync(serviceName, TimeSpan.FromSeconds(15));
                return;
            }
        }
        catch
        {
            // Fall through to command-line stop/kill.
        }

        await RunProcessAsync("sc.exe", $"stop \"{serviceName}\"");
        await WaitForServiceStoppedAsync(serviceName, TimeSpan.FromSeconds(10));

        var servicePid = await GetServiceProcessIdAsync(serviceName);
        if (servicePid > 0 && servicePid != Environment.ProcessId)
        {
            try
            {
                using var process = Process.GetProcessById(servicePid);
                if (!process.HasExited)
                {
                    process.Kill(entireProcessTree: false);
                    await process.WaitForExitAsync();
                }
            }
            catch
            {
                // Best effort only.
            }
        }
    }

    /// <summary>
    /// Waits for a service to stop without relying on ServiceController.WaitForStatus, which can throw during deletion races.
    /// </summary>
    private static async Task WaitForServiceStoppedAsync(string serviceName, TimeSpan timeout)
    {
        var startedAt = DateTime.UtcNow;

        while (DateTime.UtcNow - startedAt < timeout)
        {
            try
            {
                using var service = new ServiceController(serviceName);
                service.Refresh();

                if (service.Status == ServiceControllerStatus.Stopped)
                {
                    return;
                }
            }
            catch (InvalidOperationException)
            {
                // Service disappeared from SCM while stopping; treat as stopped for cleanup purposes.
                return;
            }
            catch
            {
                // Best effort only. Fall back to timeout behavior.
            }

            await Task.Delay(500);
        }
    }

    /// <summary>
    /// Gets the process ID for a service using sc queryex.
    /// </summary>
    private static async Task<int> GetServiceProcessIdAsync(string serviceName)
    {
        var (success, standardOutput, _) = await RunProcessAsync("sc.exe", $"queryex \"{serviceName}\"");
        if (!success)
        {
            return 0;
        }

        foreach (var line in standardOutput.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            var trimmedLine = line.Trim();
            if (!trimmedLine.StartsWith("PID", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var parts = trimmedLine.Split(':', 2);
            if (parts.Length == 2 && int.TryParse(parts[1].Trim(), out var pid))
            {
                return pid;
            }
        }

        return 0;
    }

    /// <summary>
    /// Checks whether a service still exists in the SCM.
    /// </summary>
    private static bool IsServiceRegistered(string serviceName)
    {
        return ServiceController.GetServices()
            .Any(service => string.Equals(service.ServiceName, serviceName, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Deletes a service registry key, taking ownership if required.
    /// </summary>
    private static bool TryDeleteServiceRegistryKey(string serviceName)
    {
        const string servicesRoot = @"SYSTEM\CurrentControlSet\Services";
        var serviceKeyPath = $@"{servicesRoot}\{serviceName}";

        try
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
            if (TryDeleteRegistrySubKey(baseKey, servicesRoot, serviceName))
            {
                return true;
            }

            if (!TryGrantRegistryFullControl(baseKey, serviceKeyPath))
            {
                return false;
            }

            TryGrantRegistryFullControl(baseKey, servicesRoot);
            return TryDeleteRegistrySubKey(baseKey, servicesRoot, serviceName);
        }
        catch (Exception ex)
        {
            _logger?.Error(ex, "Failed to delete service registry key for {ServiceName}", serviceName);
            return false;
        }
    }

    /// <summary>
    /// Deletes a registry subkey tree if present.
    /// </summary>
    private static bool TryDeleteRegistrySubKey(RegistryKey baseKey, string parentPath, string keyName)
    {
        using var parentKey = baseKey.OpenSubKey(parentPath, writable: true);
        if (parentKey is null)
        {
            return true;
        }

        if (!parentKey.GetSubKeyNames().Contains(keyName, StringComparer.OrdinalIgnoreCase))
        {
            return true;
        }

        parentKey.DeleteSubKeyTree(keyName, throwOnMissingSubKey: false);
        return !parentKey.GetSubKeyNames().Contains(keyName, StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Grants the Administrators group full control to a registry key.
    /// </summary>
    private static bool TryGrantRegistryFullControl(RegistryKey baseKey, string keyPath)
    {
        try
        {
            var administratorsSid = new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null);

            using (var ownershipKey = baseKey.OpenSubKey(
                       keyPath,
                       RegistryKeyPermissionCheck.ReadWriteSubTree,
                       RegistryRights.TakeOwnership | RegistryRights.ReadKey | RegistryRights.ChangePermissions))
            {
                if (ownershipKey is null)
                {
                    return false;
                }

                var security = ownershipKey.GetAccessControl(AccessControlSections.Owner);
                security.SetOwner(administratorsSid);
                ownershipKey.SetAccessControl(security);
            }

            using var permissionKey = baseKey.OpenSubKey(
                keyPath,
                RegistryKeyPermissionCheck.ReadWriteSubTree,
                RegistryRights.ChangePermissions | RegistryRights.ReadKey | RegistryRights.FullControl);

            if (permissionKey is null)
            {
                return false;
            }

            var updatedSecurity = permissionKey.GetAccessControl();
            updatedSecurity.AddAccessRule(new RegistryAccessRule(
                administratorsSid,
                RegistryRights.FullControl,
                InheritanceFlags.ContainerInherit,
                PropagationFlags.None,
                AccessControlType.Allow));
            permissionKey.SetAccessControl(updatedSecurity);

            return true;
        }
        catch
        {
            return false;
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
        if (CanReadInteractiveKey())
        {
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey(intercept: true);
        }
        else
        {
            Console.WriteLine("No interactive console detected, exiting without key prompt.");
        }
        
        return 0;
    }
#pragma warning restore Spectre1000

    /// <summary>
    /// Waits for a key press only when an interactive console is available.
    /// </summary>
    private static void WaitForUserKeyIfInteractive()
    {
        if (!CanReadInteractiveKey())
        {
            return;
        }

        Console.ReadKey(intercept: true);
    }

    /// <summary>
    /// Determines whether the current process can safely read a key from the console.
    /// </summary>
    private static bool CanReadInteractiveKey()
    {
        try
        {
            return !Console.IsInputRedirected;
        }
        catch
        {
            return false;
        }
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
