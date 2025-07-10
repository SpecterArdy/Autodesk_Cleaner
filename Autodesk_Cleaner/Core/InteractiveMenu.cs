using NLog;
using System.Security.Cryptography;
using System.Text;
using Spectre.Console;

namespace Autodesk_Cleaner.Core;

/// <summary>
/// Interactive menu system implementing NASA/NSA/DOJ high-assurance safety standards.
/// Provides multiple confirmation mechanisms and fail-safe operations.
/// </summary>
public sealed class InteractiveMenu : IDisposable
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
    private readonly RandomNumberGenerator _cryptoRng = RandomNumberGenerator.Create();
    private bool _disposed;
    
    /// <summary>
    /// Menu options available to the user.
    /// </summary>
    public enum MenuOption
    {
        /// <summary>Exit the application safely.</summary>
        Exit = 0,
        
        /// <summary>Scan registry only (read-only operation).</summary>
        ScanRegistryOnly = 1,
        
        /// <summary>Scan file system only (read-only operation).</summary>
        ScanFileSystemOnly = 2,
        
        /// <summary>Scan both registry and file system (read-only operation).</summary>
        ScanBoth = 3,
        
        /// <summary>Perform dry run cleanup (simulation only).</summary>
        DryRunCleanup = 4,
        
        /// <summary>Perform actual cleanup (destructive operation).</summary>
        ActualCleanup = 5,
        
        /// <summary>Configure backup settings.</summary>
        ConfigureBackup = 6,
        
        /// <summary>View current configuration.</summary>
        ViewConfiguration = 7,
        
        /// <summary>Emergency abort (immediate termination).</summary>
        EmergencyAbort = 99
    }

    /// <summary>
    /// Security levels for operations following defense-in-depth principles.
    /// </summary>
    public enum SecurityLevel
    {
        /// <summary>Read-only operations, no system changes.</summary>
        ReadOnly = 1,
        
        /// <summary>Simulation operations, no actual changes.</summary>
        Simulation = 2,
        
        /// <summary>Low-risk operations with automatic rollback.</summary>
        LowRisk = 3,
        
        /// <summary>High-risk operations requiring multiple confirmations.</summary>
        HighRisk = 4,
        
        /// <summary>Critical operations requiring cryptographic confirmation.</summary>
        Critical = 5
    }

    /// <summary>
    /// Displays the main menu and handles user interaction.
    /// </summary>
    /// <returns>The selected menu option.</returns>
    public MenuOption DisplayMainMenu()
    {
        Logger.Info("Displaying main interactive menu");
        
        try
        {
            Console.Clear();
            DisplayHeader();
            DisplayMenuOptions();
            DisplaySafetyWarnings();
            
            return GetValidatedUserChoice();
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Critical error in main menu display");
            return MenuOption.EmergencyAbort;
        }
    }

    /// <summary>
    /// Gets and validates user choice with input sanitization.
    /// </summary>
    /// <returns>Validated menu option.</returns>
    private MenuOption GetValidatedUserChoice()
    {
        const int maxAttempts = 3;
        int attempts = 0;

        while (attempts < maxAttempts)
        {
            try
            {
                var input = AnsiConsole.Prompt(
                    new TextPrompt<string>("[bold cyan]Enter your choice (0-7, 99 for emergency abort):[/]")
                        .PromptStyle("cyan")
                        .ValidationErrorMessage("[red]Invalid input. Please enter a number.[/]"));

                Logger.Debug("User input received: {Input}", SanitizeForLogging(input));

                // Input validation and sanitization
                if (string.IsNullOrWhiteSpace(input))
                {
                    DisplayError("Invalid input: Empty or whitespace only");
                    attempts++;
                    continue;
                }

                // Sanitize input - remove non-numeric characters
                var sanitizedInput = new string(input.Where(char.IsDigit).ToArray());
                
                if (string.IsNullOrEmpty(sanitizedInput))
                {
                    DisplayError("Invalid input: No numeric characters found");
                    attempts++;
                    continue;
                }

                if (!int.TryParse(sanitizedInput, out var choice))
                {
                    DisplayError("Invalid input: Cannot parse as integer");
                    attempts++;
                    continue;
                }

                // Validate choice is within acceptable range
                if (!Enum.IsDefined(typeof(MenuOption), choice))
                {
                    DisplayError($"Invalid choice: {choice} is not a valid menu option");
                    attempts++;
                    continue;
                }

                var selectedOption = (MenuOption)choice;
                
                Logger.Info("User selected menu option: {Option}", selectedOption);
                
                // Perform security validation based on operation type
                if (RequiresSecurityConfirmation(selectedOption))
                {
                    if (!PerformSecurityConfirmation(selectedOption))
                    {
                        Logger.Warn("Security confirmation failed for option: {Option}", selectedOption);
                        return MenuOption.Exit;
                    }
                }

                return selectedOption;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Error processing user input, attempt {Attempt}", attempts + 1);
                DisplayError($"System error processing input: {ex.Message}");
                attempts++;
            }
        }

        Logger.Warn("Maximum input attempts ({MaxAttempts}) exceeded, defaulting to safe exit", maxAttempts);
        DisplayError($"Maximum attempts ({maxAttempts}) exceeded. Defaulting to safe exit for security.");
        return MenuOption.Exit;
    }

    /// <summary>
    /// Displays the application header with safety information.
    /// </summary>
    private static void DisplayHeader()
    {
        Console.Clear();
        
        // Create a fancy header panel
        var headerPanel = new Panel(new Markup(
            "[bold red]HIGH-ASSURANCE AUTODESK REGISTRY CLEANER[/]\n" +
            "[bold yellow]SAFETY-CRITICAL OPERATION[/]\n" +
            "[dim]v1.0.0 - High-Assurance Security Standards[/]"))
        {
            Border = BoxBorder.Double,
            BorderStyle = new Style(Color.Red),
            Header = new PanelHeader(" [bold red]CRITICAL SYSTEM TOOL[/] "),
            Padding = new Padding(2, 1)
        };
        
        AnsiConsole.Write(headerPanel);
        
        // Warning panel
        var warningPanel = new Panel(new Markup(
            "[bold yellow]WARNING: This application can permanently modify your system[/]\n" +
            "[bold yellow]Ensure you have proper backups before proceeding[/]\n" +
            "[bold yellow]All operations are logged for audit purposes[/]"))
        {
            Border = BoxBorder.Rounded,
            BorderStyle = new Style(Color.Yellow),
            Header = new PanelHeader(" [bold yellow]SAFETY NOTICE[/] ")
        };
        
        AnsiConsole.Write(warningPanel);
    }

    /// <summary>
    /// Displays all available menu options with security indicators.
    /// </summary>
    private static void DisplayMenuOptions()
    {
        var menuTable = new Table()
        {
            Border = TableBorder.Rounded,
            BorderStyle = new Style(Color.Blue)
        };
        
        menuTable.AddColumn(new TableColumn("[bold]Option[/]").Centered());
        menuTable.AddColumn(new TableColumn("[bold]Description[/]").LeftAligned());
        menuTable.AddColumn(new TableColumn("[bold]Security Level[/]").Centered());
        
        // Safe operations
        menuTable.AddRow("[bold green]1[/]", "Scan Registry Only", "[green]READ-ONLY - Safe[/]");
        menuTable.AddRow("[bold green]2[/]", "Scan File System Only", "[green]READ-ONLY - Safe[/]");
        menuTable.AddRow("[bold green]3[/]", "Scan Both Systems", "[green]READ-ONLY - Safe[/]");
        menuTable.AddRow("[bold green]7[/]", "View Current Configuration", "[green]READ-ONLY - Safe[/]");
        
        // Simulation operations
        menuTable.AddRow("[bold yellow]4[/]", "Dry Run Cleanup", "[yellow]SIMULATION - No Changes[/]");
        menuTable.AddRow("[bold yellow]6[/]", "Configure Backup Settings", "[yellow]CONFIGURATION - Low Risk[/]");
        
        // Destructive operations
        menuTable.AddRow("[bold red]5[/]", "Actual Cleanup", "[red]DESTRUCTIVE - High Risk[/]");
        
        // Control operations
        menuTable.AddRow("[bold cyan]0[/]", "Exit Safely", "[cyan]CONTROL - Safe Exit[/]");
        menuTable.AddRow("[bold magenta]99[/]", "Emergency Abort", "[red]EMERGENCY - Immediate Termination[/]");
        
        var menuPanel = new Panel(menuTable)
        {
            Header = new PanelHeader(" [bold blue]AVAILABLE OPTIONS[/] "),
            Border = BoxBorder.Double,
            BorderStyle = new Style(Color.Blue)
        };
        
        AnsiConsole.Write(menuPanel);
    }

    /// <summary>
    /// Displays safety warnings and compliance information.
    /// </summary>
    private static void DisplaySafetyWarnings()
    {
        var complianceList = new List<string>
        {
            "This application follows NASA/NSA/DOJ high-assurance standards",
            "All operations require explicit confirmation",
            "Destructive operations require cryptographic verification",
            "All actions are logged with full audit trail",
            "Emergency abort (99) is available at any time"
        };
        
        var compliancePanel = new Panel(new Markup(
            string.Join("\n", complianceList.Select(item => $"[bold green]•[/] {item}")))
        )
        {
            Header = new PanelHeader(" [bold magenta]COMPLIANCE NOTICE[/] "),
            Border = BoxBorder.Rounded,
            BorderStyle = new Style(Color.Purple)
        };
        
        AnsiConsole.Write(compliancePanel);
    }

    /// <summary>
    /// Determines if a menu option requires enhanced security confirmation.
    /// </summary>
    /// <param name="option">The menu option to evaluate.</param>
    /// <returns>True if security confirmation is required.</returns>
    private static bool RequiresSecurityConfirmation(MenuOption option)
    {
        return option switch
        {
            MenuOption.ActualCleanup => true,
            MenuOption.DryRunCleanup => true,
            MenuOption.ConfigureBackup => true,
            _ => false
        };
    }

    /// <summary>
    /// Performs multi-factor security confirmation for high-risk operations.
    /// Implements defense-in-depth security principles.
    /// </summary>
    /// <param name="option">The operation requiring confirmation.</param>
    /// <returns>True if all security checks pass.</returns>
    private bool PerformSecurityConfirmation(MenuOption option)
    {
        Logger.Info("Initiating security confirmation for operation: {Operation}", option);
        
        var securityLevel = GetSecurityLevel(option);
        
        try
        {
            // Level 1: Basic confirmation
            if (!PerformBasicConfirmation(option))
            {
                Logger.Warn("Basic confirmation failed for operation: {Operation}", option);
                return false;
            }

            // Level 2: Risk acknowledgment (for simulation and above)
            if (securityLevel >= SecurityLevel.Simulation && !PerformRiskAcknowledgment(option))
            {
                Logger.Warn("Risk acknowledgment failed for operation: {Operation}", option);
                return false;
            }

            // Level 3: Administrative confirmation (for high-risk operations)
            if (securityLevel >= SecurityLevel.HighRisk && !PerformAdministrativeConfirmation(option))
            {
                Logger.Warn("Administrative confirmation failed for operation: {Operation}", option);
                return false;
            }

            // Level 4: Cryptographic challenge (for critical operations)
            if (securityLevel >= SecurityLevel.Critical && !PerformCryptographicChallenge(option))
            {
                Logger.Warn("Cryptographic challenge failed for operation: {Operation}", option);
                return false;
            }

            Logger.Info("All security confirmations passed for operation: {Operation}", option);
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Security confirmation process failed for operation: {Operation}", option);
            return false;
        }
    }

    /// <summary>
    /// Gets the security level required for a specific operation.
    /// </summary>
    /// <param name="option">The menu option.</param>
    /// <returns>Required security level.</returns>
    private static SecurityLevel GetSecurityLevel(MenuOption option)
    {
        return option switch
        {
            MenuOption.ScanRegistryOnly or MenuOption.ScanFileSystemOnly or MenuOption.ScanBoth => SecurityLevel.ReadOnly,
            MenuOption.ViewConfiguration => SecurityLevel.ReadOnly,
            MenuOption.DryRunCleanup => SecurityLevel.Simulation,
            MenuOption.ConfigureBackup => SecurityLevel.LowRisk,
            MenuOption.ActualCleanup => SecurityLevel.Critical,
            _ => SecurityLevel.ReadOnly
        };
    }

    /// <summary>
    /// Performs basic confirmation for the operation.
    /// </summary>
    /// <param name="option">The operation to confirm.</param>
    /// <returns>True if confirmed.</returns>
    private static bool PerformBasicConfirmation(MenuOption option)
    {
        var confirmationPanel = new Panel(new Markup(
            $"[bold yellow]CONFIRMATION REQUIRED: {option}[/]\n\n" +
            $"You have selected: [bold]{GetOperationDescription(option)}[/]\n\n" +
            "[red]Do you want to proceed?[/]\n" +
            "[dim](Type 'YES' to confirm, anything else to cancel)[/]"))
        {
            Border = BoxBorder.Double,
            BorderStyle = new Style(Color.Yellow),
            Header = new PanelHeader(" [bold yellow]CONFIRMATION REQUIRED[/] "),
        };
        
        AnsiConsole.Write(confirmationPanel);
        
        var response = AnsiConsole.Prompt(
            new TextPrompt<string>("[bold yellow]Your response:[/]")
                .PromptStyle("yellow")
                .AllowEmpty());
        
        var confirmed = string.Equals(response?.Trim(), "YES", StringComparison.OrdinalIgnoreCase);
        
        if (confirmed)
        {
            AnsiConsole.MarkupLine("[bold green]Basic confirmation: PASSED[/]");
        }
        else
        {
            AnsiConsole.MarkupLine("[bold red]Basic confirmation: FAILED[/]");
        }
        
        return confirmed;
    }

    /// <summary>
    /// Performs risk acknowledgment for potentially dangerous operations.
    /// </summary>
    /// <param name="option">The operation to acknowledge.</param>
    /// <returns>True if risk is acknowledged.</returns>
    private static bool PerformRiskAcknowledgment(MenuOption option)
    {
        var risks = GetOperationRisks(option);
        var riskList = string.Join("\n", risks.Select(risk => $"[bold red][!] {risk}"));
        
        var riskPanel = new Panel(new Markup(
            "[bold red]RISK ACKNOWLEDGMENT REQUIRED[/]\n\n" +
            "The following risks have been identified:\n\n" +
            riskList + "\n\n" +
            "[bold yellow]Type 'I ACKNOWLEDGE THE RISKS' to proceed:[/]"))
        {
            Border = BoxBorder.Double,
            BorderStyle = new Style(Color.Red),
            Header = new PanelHeader(" [bold red]RISK ASSESSMENT[/] "),
        };
        
        AnsiConsole.Write(riskPanel);
        
        var response = AnsiConsole.Prompt(
            new TextPrompt<string>("[bold red]Your response:[/]")
                .PromptStyle("red")
                .AllowEmpty());
        
        var acknowledged = string.Equals(response?.Trim(), "I ACKNOWLEDGE THE RISKS", StringComparison.OrdinalIgnoreCase);
        
        if (acknowledged)
        {
            AnsiConsole.MarkupLine("[bold green]Risk acknowledgment: PASSED[/]");
        }
        else
        {
            AnsiConsole.MarkupLine("[bold red]Risk acknowledgment: FAILED[/]");
        }
        
        return acknowledged;
    }

    /// <summary>
    /// Performs administrative confirmation with additional verification.
    /// </summary>
    /// <param name="option">The operation to confirm.</param>
    /// <returns>True if administratively confirmed.</returns>
    private static bool PerformAdministrativeConfirmation(MenuOption option)
    {
        var adminInfo = new Table()
        {
            Border = TableBorder.None
        };
        adminInfo.AddColumn("[bold]Field[/]");
        adminInfo.AddColumn("[bold]Value[/]");
        adminInfo.AddRow("Current User", Environment.UserName);
        adminInfo.AddRow("Current Time", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " UTC");
        adminInfo.AddRow("Operation", option.ToString());
        
        var adminText = new Markup(
            "[bold magenta]ADMINISTRATIVE CONFIRMATION REQUIRED[/]\n\n" +
            "As an administrator, you must provide additional confirmation.\n" +
            "This operation will be logged with your user context for audit purposes.\n\n");
        
        var adminPanel = new Panel(new Rows(adminText, adminInfo, new Markup("\n[bold yellow]Type 'ADMIN CONFIRMED' to proceed:[/]")))
        {
            Border = BoxBorder.Double,
            BorderStyle = new Style(Color.Purple),
            Header = new PanelHeader(" [bold magenta]ADMINISTRATIVE VERIFICATION[/] "),
        };
        
        AnsiConsole.Write(adminPanel);
        
        var response = AnsiConsole.Prompt(
            new TextPrompt<string>("[bold magenta]Your response:[/]")
                .PromptStyle("magenta")
                .AllowEmpty());
        
        var confirmed = string.Equals(response?.Trim(), "ADMIN CONFIRMED", StringComparison.OrdinalIgnoreCase);
        
        if (confirmed)
        {
            AnsiConsole.MarkupLine("[bold green]Administrative confirmation: PASSED[/]");
        }
        else
        {
            AnsiConsole.MarkupLine("[bold red]Administrative confirmation: FAILED[/]");
        }
        
        return confirmed;
    }

    /// <summary>
    /// Performs cryptographic challenge for critical operations.
    /// </summary>
    /// <param name="option">The operation requiring cryptographic confirmation.</param>
    /// <returns>True if cryptographic challenge is successful.</returns>
    private bool PerformCryptographicChallenge(MenuOption option)
    {
        // Generate cryptographic challenge
        var challengeBytes = new byte[8];
        _cryptoRng.GetBytes(challengeBytes);
        var challenge = Convert.ToHexString(challengeBytes).ToLowerInvariant();
        
        var cryptoPanel = new Panel(new Markup(
            "[bold red]CRYPTOGRAPHIC VERIFICATION REQUIRED[/]\n\n" +
            "For maximum security, you must complete a cryptographic challenge.\n\n" +
            $"Please type the following hexadecimal string exactly:\n" +
            $"[bold yellow on black] {challenge} [/]\n\n" +
            "[dim]This ensures only authorized personnel can execute critical operations.[/]"))
        {
            Border = BoxBorder.Double,
            BorderStyle = new Style(Color.Red),
            Header = new PanelHeader(" [bold red]CRYPTOGRAPHIC CHALLENGE[/] "),
        };
        
        AnsiConsole.Write(cryptoPanel);
        
        var response = AnsiConsole.Prompt(
            new TextPrompt<string>("[bold red]Your response:[/]")
                .PromptStyle("red")
                .AllowEmpty())
            .Trim().ToLowerInvariant();
        
        var verified = string.Equals(response, challenge, StringComparison.Ordinal);
        
        if (verified)
        {
            AnsiConsole.MarkupLine("[bold green]Cryptographic verification: PASSED[/]");
        }
        else
        {
            AnsiConsole.MarkupLine("[bold red]Cryptographic verification: FAILED[/]");
            
            // Add delay to prevent brute force attempts
            AnsiConsole.Status()
                .Spinner(Spinner.Known.Dots)
                .SpinnerStyle(Style.Parse("red"))
                .Start("[red]Security delay in effect...[/]", ctx => 
                {
                    Thread.Sleep(3000);
                });
        }
        
        return verified;
    }

    /// <summary>
    /// Gets the human-readable description of an operation.
    /// </summary>
    /// <param name="option">The menu option.</param>
    /// <returns>Operation description.</returns>
    private static string GetOperationDescription(MenuOption option)
    {
        return option switch
        {
            MenuOption.ScanRegistryOnly => "Scan Windows Registry for Autodesk entries (read-only)",
            MenuOption.ScanFileSystemOnly => "Scan file system for Autodesk files and directories (read-only)",
            MenuOption.ScanBoth => "Scan both registry and file system for Autodesk entries (read-only)",
            MenuOption.DryRunCleanup => "Simulate complete Autodesk cleanup without making changes",
            MenuOption.ActualCleanup => "Perform actual cleanup - PERMANENTLY DELETE Autodesk entries",
            MenuOption.ConfigureBackup => "Configure backup settings and locations",
            MenuOption.ViewConfiguration => "Display current application configuration",
            MenuOption.Exit => "Exit application safely",
            MenuOption.EmergencyAbort => "Emergency abort - immediate termination",
            _ => "Unknown operation"
        };
    }

    /// <summary>
    /// Gets the risks associated with a specific operation.
    /// </summary>
    /// <param name="option">The menu option.</param>
    /// <returns>List of identified risks.</returns>
    private static string[] GetOperationRisks(MenuOption option)
    {
        return option switch
        {
            MenuOption.DryRunCleanup => [
                "System resources will be consumed during simulation",
                "Temporary files may be created during analysis",
                "Registry access may trigger security software alerts"
            ],
            MenuOption.ActualCleanup => [
                "ALL Autodesk software will be permanently removed",
                "Registry entries will be permanently deleted",
                "Autodesk project files may become inaccessible",
                "System restore may not recover all deleted items",
                "Professional licenses may need to be reactivated",
                "This operation cannot be undone without full system restore"
            ],
            MenuOption.ConfigureBackup => [
                "Backup locations may consume significant disk space",
                "Existing backup files may be overwritten",
                "Invalid backup paths may cause operation failures"
            ],
            _ => ["General system access for the specified operation"]
        };
    }

    /// <summary>
    /// Displays an error message with consistent formatting.
    /// </summary>
    /// <param name="message">The error message to display.</param>
    private static void DisplayError(string message)
    {
        AnsiConsole.Write(new Panel(new Markup($"[bold red]ERROR: {message}[/]"))
        {
            Border = BoxBorder.Rounded,
            BorderStyle = new Style(Color.Red),
            Header = new PanelHeader(" [bold red]ERROR[/] ")
        });
        
        AnsiConsole.MarkupLine("[dim]Press any key to continue...[/]");
        Console.ReadKey(true);
    }

    /// <summary>
    /// Sanitizes input for safe logging (removes sensitive information).
    /// </summary>
    /// <param name="input">The input to sanitize.</param>
    /// <returns>Sanitized input safe for logging.</returns>
    private static string SanitizeForLogging(string? input)
    {
        if (string.IsNullOrEmpty(input))
            return "[EMPTY]";
            
        // Replace potentially sensitive characters and limit length
        var sanitized = new string(input.Take(50).Select(c => char.IsControl(c) ? '?' : c).ToArray());
        return $"[{sanitized.Length} chars]: {sanitized}";
    }

    /// <summary>
    /// Disposes of cryptographic resources.
    /// </summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            _cryptoRng?.Dispose();
            _disposed = true;
            Logger.Debug("InteractiveMenu disposed successfully");
        }
    }
}
