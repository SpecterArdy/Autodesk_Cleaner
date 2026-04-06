using Microsoft.Win32;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Diagnostics;
using System.Text.RegularExpressions;
using NLog;
using System.IO.Abstractions;
using Spectre.Console;

namespace Autodesk_Cleaner.Core;

/// <summary>
/// Implementation of IRegistryScanner specifically designed for Autodesk product cleanup.
/// </summary>
public sealed class AutodeskRegistryScanner : IRegistryScanner, IDisposable
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
    private readonly ScannerConfig _config;
    private readonly List<RegistryEntry> _foundEntries = [];
    private readonly List<string> _errors = [];
    private bool _disposed;

    /// <summary>
    /// Registry paths to scan for Autodesk entries.
    /// </summary>
    private static readonly IReadOnlyList<string> AutodeskRegistryPaths = [
        @"SOFTWARE\Autodesk",
        @"SOFTWARE\WOW6432Node\Autodesk",
        @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        @"SOFTWARE\Classes\Installer\Products",
        @"SOFTWARE\Classes\Installer\Features",
        @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"
    ];

    /// <summary>
    /// Patterns to match Autodesk-related registry entries.
    /// </summary>
    private static readonly Regex AutodeskPattern = new(
        @"(autodesk|maya|3ds\s*max|autocad|revit|inventor|fusion|vault|navisworks|mudbox|motionbuilder|alias|adsk)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Initializes a new instance of the AutodeskRegistryScanner.
    /// </summary>
    /// <param name="config">Configuration options for the scanner.</param>
    public AutodeskRegistryScanner(ScannerConfig config = default)
    {
        _config = config == default ? ScannerConfig.Default : config;
        Logger.Info("AutodeskRegistryScanner initialized with config: {@Config}", _config);
    }

    /// <inheritdoc />
    public async Task<IReadOnlyCollection<RegistryEntry>> ScanRegistryAsync()
    {
        Logger.Info("Starting registry scan with configuration: {@Config}", _config);
        _foundEntries.Clear();
        _errors.Clear();

        try
        {
            // Scan HKEY_LOCAL_MACHINE if enabled
            if (_config.IncludeLocalMachine)
            {
                Logger.Info("Scanning HKEY_LOCAL_MACHINE");
                await ScanRegistryHiveAsync(Registry.LocalMachine, RegistryHive.LocalMachine);
                Logger.Info("Completed HKEY_LOCAL_MACHINE scan, found {EntryCount} entries", _foundEntries.Count);
            }

            // Scan HKEY_CURRENT_USER if enabled
            if (_config.IncludeUserHive)
            {
                Logger.Info("Scanning HKEY_CURRENT_USER");
                var initialCount = _foundEntries.Count;
                await ScanRegistryHiveAsync(Registry.CurrentUser, RegistryHive.CurrentUser);
                Logger.Info("Completed HKEY_CURRENT_USER scan, found {NewEntryCount} additional entries", _foundEntries.Count - initialCount);
            }

            Logger.Info("Registry scan completed. Total entries found: {TotalEntries}, Errors: {ErrorCount}", _foundEntries.Count, _errors.Count);
            // Return a copy to avoid issues with disposal clearing the internal list
            return _foundEntries.ToList().AsReadOnly();
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Critical error during registry scan");
            _errors.Add($"Critical error during registry scan: {ex.Message}");
            throw new InvalidOperationException("Registry scan failed", ex);
        }
    }

    /// <inheritdoc />
    public async Task<RemovalResult> RemoveEntriesAsync(IReadOnlyCollection<RegistryEntry> entries)
    {
        // Filter to only valid entries
        var validEntries = FilterValidEntries(entries);
        var removalPlan = BuildRemovalPlan(validEntries);
        
        if (removalPlan.Count == 0)
        {
            Logger.Error("Registry entry validation failed - no valid entries found to process");
            return new RemovalResult(
                TotalEntries: 0,
                SuccessfulRemovals: 0,
                FailedRemovals: 0,
                Errors: ["Entry validation failed - no Autodesk-related entries found to process"],
                Duration: TimeSpan.Zero);
        }
        
        Logger.Info("Processing {RemovalPlanCount} planned registry removals from {ValidCount} valid entries out of {TotalCount} total entries",
            removalPlan.Count, validEntries.Count, entries.Count);

        var stopwatch = Stopwatch.StartNew();
        var errors = new List<string>();
        var successfulRemovals = 0;

        // Create backup if requested
        if (_config.CreateBackup)
        {
            var backupPath = Path.Combine(_config.BackupPath, $"registry_backup_{DateTime.Now:yyyyMMdd_HHmmss}.reg");
            Logger.Info("Creating registry backup at: {BackupPath}", backupPath);
            if (!await CreateBackupAsync(backupPath))
            {
                Logger.Warn("Failed to create registry backup at: {BackupPath}; continuing without registry backup", backupPath);
                errors.Add("Failed to create registry backup");
            }
            else
            {
                Logger.Info("Registry backup created successfully at: {BackupPath}", backupPath);
            }
        }

        // Stop Autodesk services before registry modifications
        await StopAutodeskServicesAsync();

        // Process each valid entry
        if (!removalPlan.Any())
        {
            Logger.Warn("No valid registry entries to process.");
            return new RemovalResult(
                TotalEntries: 0,
                SuccessfulRemovals: 0,
                FailedRemovals: 0,
                Errors: ["No valid entries found"],
                Duration: TimeSpan.Zero
            );
        }

        foreach (var entry in removalPlan)
        {
            try
            {
                if (_config.DryRun)
                {
                    Logger.Debug("[DRY RUN] Would remove registry entry: {EntryPath}", entry.DisplayName);
                    AnsiConsole.MarkupLine($"[dim][DRY RUN] Would remove: {entry.DisplayName}[/]");
                    successfulRemovals++;
                    continue;
                }

                Logger.Debug("Removing registry entry: {EntryPath}", entry.DisplayName);
                if (await RemoveRegistryEntryAsync(entry))
                {
                    Logger.Debug("Successfully removed registry entry: {EntryPath}", entry.DisplayName);
                    successfulRemovals++;
                }
                else
                {
                    Logger.Warn("Failed to remove registry entry: {EntryPath}", entry.DisplayName);
                    errors.Add($"Failed to remove: {entry.DisplayName}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Error removing registry entry {EntryPath}: {Message}", entry.DisplayName, ex.Message);
                errors.Add($"Error removing {entry.DisplayName}: {ex.Message}");
            }
        }

        stopwatch.Stop();

        return new RemovalResult(
            TotalEntries: removalPlan.Count,
            SuccessfulRemovals: successfulRemovals,
            FailedRemovals: removalPlan.Count - successfulRemovals,
            Errors: errors,
            Duration: stopwatch.Elapsed);
    }

    /// <inheritdoc />
    public async Task<bool> CreateBackupAsync(string backupPath)
    {
        try
        {
            var directory = Path.GetDirectoryName(backupPath);
            if (!string.IsNullOrEmpty(directory))
            {
                Logger.Debug("Creating backup directory: {Directory}", directory);
                Directory.CreateDirectory(directory);
                Logger.Debug("Backup directory created successfully: {Directory}", directory);
            }

            // Use reg.exe to export registry
            var startInfo = new ProcessStartInfo
            {
                FileName = "reg.exe",
                Arguments = $"export \"HKEY_LOCAL_MACHINE\\SOFTWARE\\Autodesk\" \"{backupPath}\" /y",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using var process = Process.Start(startInfo);
            if (process is not null)
            {
                await process.WaitForExitAsync();
                return process.ExitCode == 0;
            }

            return false;
        }
        catch (Exception ex)
        {
            _errors.Add($"Backup creation failed: {ex.Message}");
            return false;
        }
    }

    /// <inheritdoc />
    public bool ValidateEntries(IReadOnlyCollection<RegistryEntry> entries)
    {
        if (entries is null || entries.Count == 0)
        {
            Logger.Warn("ValidateEntries failed: entries is null or empty");
            return false;
        }

        Logger.Info("Validating {EntryCount} registry entries", entries.Count);
        AnsiConsole.WriteLine($"[DEBUG] Starting validation of {entries.Count} registry entries");
        
        // Check each entry and track valid/invalid for debugging
        var validEntries = new List<RegistryEntry>();
        var invalidEntries = new List<RegistryEntry>();
        
        foreach (var entry in entries)
        {
            var isAutodeskPath = IsAutodeskRegistryPath(entry.KeyPath);
            var matchesKeyPath = AutodeskPattern.IsMatch(entry.KeyPath);
            var matchesValueName = entry.ValueName is not null && AutodeskPattern.IsMatch(entry.ValueName);
            var matchesValueData = entry.ValueData is string valueData && AutodeskPattern.IsMatch(valueData);
            
            var isValid = isAutodeskPath || matchesKeyPath || matchesValueName || matchesValueData;
            
            if (isValid)
            {
                validEntries.Add(entry);
            }
            else
            {
                invalidEntries.Add(entry);
                Logger.Info("Invalid entry: {KeyPath}, AutodeskPath: {IsAutodeskPath}, KeyMatch: {MatchesKeyPath}, ValueNameMatch: {MatchesValueName}, ValueDataMatch: {MatchesValueData}", 
                    entry.KeyPath, isAutodeskPath, matchesKeyPath, matchesValueName, matchesValueData);
            }
        }
        
        Logger.Info("Validation results: {ValidCount} valid entries, {InvalidCount} invalid entries out of {TotalCount} total", 
            validEntries.Count, invalidEntries.Count, entries.Count);
        AnsiConsole.WriteLine($"[DEBUG] Found {validEntries.Count} valid entries and {invalidEntries.Count} invalid entries");
        
        if (invalidEntries.Count > 0)
        {
            Logger.Warn("Found {InvalidCount} invalid entries out of {TotalCount} - these will be filtered out", invalidEntries.Count, entries.Count);
            // Log first few invalid entries for debugging
            foreach (var entry in invalidEntries.Take(5))
            {
                Logger.Debug("Invalid entry example: {KeyPath} (ValueName: {ValueName})", entry.KeyPath, entry.ValueName ?? "<null>");
                AnsiConsole.WriteLine($"[DEBUG] Invalid entry: {entry.KeyPath}");
            }
        }
        
        // Return true if we have ANY valid entries to process
        var hasValidEntries = validEntries.Count > 0;
        Logger.Info("Validation result: {HasValidEntries} - {ValidCount} entries will be processed", 
            hasValidEntries ? "PASS" : "FAIL", validEntries.Count);
        AnsiConsole.WriteLine($"[DEBUG] Validation result: {(hasValidEntries ? "PASS" : "FAIL")}");
        
        return hasValidEntries;
    }

    /// <summary>
    /// Gets the count of valid Autodesk-related entries from a collection.
    /// </summary>
    /// <param name="entries">The registry entries to analyze.</param>
    /// <returns>The number of valid entries.</returns>
    public int GetValidEntryCount(IReadOnlyCollection<RegistryEntry> entries)
    {
        Logger.Info("GetValidEntryCount: Starting count for {EntryCount} registry entries", entries?.Count ?? 0);
        if (entries == null || entries.Count == 0)
        {
            Logger.Warn("GetValidEntryCount: Null or empty entries provided");
            return 0;
        }
        var validEntries = FilterValidEntries(entries);
        Logger.Info("GetValidEntryCount: Found {ValidCount} valid registry entries", validEntries.Count);
        return validEntries.Count;
    }

    /// <summary>
    /// Filters registry entries to only include valid Autodesk-related entries.
    /// </summary>
    /// <param name="entries">The registry entries to filter.</param>
    /// <returns>A list of valid Autodesk-related entries.</returns>
    private List<RegistryEntry> FilterValidEntries(IReadOnlyCollection<RegistryEntry> entries)
    {
        if (entries is null || entries.Count == 0)
        {
            Logger.Info("FilterValidEntries: entries is null or empty");
            return new List<RegistryEntry>();
        }

        Logger.Info("FilterValidEntries: Processing {EntryCount} entries", entries.Count);
        var validEntries = new List<RegistryEntry>();
        var invalidCount = 0;
        
        foreach (var entry in entries)
        {
            var isAutodeskPath = IsAutodeskRegistryPath(entry.KeyPath);
            var matchesKeyPath = AutodeskPattern.IsMatch(entry.KeyPath);
            var matchesValueName = entry.ValueName is not null && AutodeskPattern.IsMatch(entry.ValueName);
            var matchesValueData = entry.ValueData is string valueData && AutodeskPattern.IsMatch(valueData);
            
            var isValid = isAutodeskPath || matchesKeyPath || matchesValueName || matchesValueData;
            
            if (isValid)
            {
                validEntries.Add(entry);
            }
            else
            {
                invalidCount++;
                if (invalidCount <= 3) // Log first 3 invalid entries for debugging
                {
                    Logger.Debug("FilterValidEntries: Invalid entry {KeyPath} - AutodeskPath: {IsAutodeskPath}, KeyMatch: {MatchesKeyPath}, ValueNameMatch: {MatchesValueName}, ValueDataMatch: {MatchesValueData}", 
                        entry.KeyPath, isAutodeskPath, matchesKeyPath, matchesValueName, matchesValueData);
                }
            }
        }
        
        Logger.Info("FilterValidEntries: Found {ValidCount} valid entries, {InvalidCount} invalid entries", validEntries.Count, invalidCount);
        return validEntries;
    }

    /// <summary>
    /// Builds the smallest safe registry removal plan.
    /// Prefers deleting a parent Autodesk key once instead of deleting each child key/value separately.
    /// </summary>
    /// <param name="entries">Validated Autodesk registry entries.</param>
    /// <returns>The optimized removal plan.</returns>
    private static List<RegistryEntry> BuildRemovalPlan(IReadOnlyCollection<RegistryEntry> entries)
    {
        if (entries.Count == 0)
        {
            return [];
        }

        var orderedEntries = entries
            .OrderBy(entry => entry.EntryType == RegistryEntryType.Key ? 0 : 1)
            .ThenBy(entry => entry.KeyPath.Length)
            .ThenBy(entry => entry.KeyPath, StringComparer.OrdinalIgnoreCase)
            .ThenBy(entry => entry.ValueName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var plannedEntries = new List<RegistryEntry>();

        foreach (var entry in orderedEntries)
        {
            var coveredByParentKey = plannedEntries.Any(plannedEntry =>
                plannedEntry.RegistryHive == entry.RegistryHive &&
                plannedEntry.EntryType == RegistryEntryType.Key &&
                IsSameOrDescendantRegistryPath(entry.KeyPath, plannedEntry.KeyPath));

            if (!coveredByParentKey)
            {
                plannedEntries.Add(entry);
            }
        }

        return plannedEntries;
    }

    /// <summary>
    /// Checks if a registry path is under a known Autodesk registry location.
    /// </summary>
    /// <param name="keyPath">The registry key path to check.</param>
    /// <returns>True if the path is under an Autodesk registry location.</returns>
    private static bool IsAutodeskRegistryPath(string keyPath)
    {
        if (string.IsNullOrEmpty(keyPath))
            return false;

        return AutodeskRegistryPaths.Any(basePath => 
            keyPath.StartsWith(basePath, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Scans a specific registry hive for Autodesk entries.
    /// </summary>
    /// <param name="registryKey">The registry key to scan.</param>
    /// <param name="hive">The registry hive being scanned.</param>
    private async Task ScanRegistryHiveAsync(RegistryKey registryKey, RegistryHive hive)
    {
        foreach (var basePath in AutodeskRegistryPaths)
        {
            try
            {
                using var key = registryKey.OpenSubKey(basePath, false);
                if (key is not null)
                {
                    await ScanRegistryKeyAsync(key, basePath, hive, 0);
                }
            }
            catch (Exception ex)
            {
                _errors.Add($"Error scanning {hive}\\{basePath}: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Recursively scans a registry key for Autodesk entries.
    /// </summary>
    /// <param name="key">The registry key to scan.</param>
    /// <param name="keyPath">The path to the registry key.</param>
    /// <param name="hive">The registry hive being scanned.</param>
    /// <param name="depth">Current depth in the registry tree.</param>
    private async Task ScanRegistryKeyAsync(RegistryKey key, string keyPath, RegistryHive hive, int depth)
    {
        if (depth > _config.MaxDepth)
        {
            return;
        }

        try
        {
            // Check if the key itself is Autodesk-related
            if (AutodeskPattern.IsMatch(keyPath))
            {
                _foundEntries.Add(new RegistryEntry(
                    KeyPath: keyPath,
                    ValueName: null,
                    ValueData: null,
                    RegistryHive: hive,
                    EntryType: RegistryEntryType.Key,
                    LastWriteTime: DateTime.Now));
            }

            // Scan values
            foreach (var valueName in key.GetValueNames())
            {
                try
                {
                    var valueData = key.GetValue(valueName);
                    var valueDataString = valueData?.ToString() ?? string.Empty;

                    if (AutodeskPattern.IsMatch(valueName) || 
                        AutodeskPattern.IsMatch(valueDataString))
                    {
                        _foundEntries.Add(new RegistryEntry(
                            KeyPath: keyPath,
                            ValueName: valueName,
                            ValueData: valueData,
                            RegistryHive: hive,
                            EntryType: RegistryEntryType.Value,
                            LastWriteTime: DateTime.Now));
                    }
                }
                catch (Exception ex)
                {
                    _errors.Add($"Error reading value {valueName} in {keyPath}: {ex.Message}");
                }
            }

            // Scan subkeys
            foreach (var subKeyName in key.GetSubKeyNames())
            {
                try
                {
                    using var subKey = key.OpenSubKey(subKeyName, false);
                    if (subKey is not null)
                    {
                        var subKeyPath = $"{keyPath}\\{subKeyName}";
                        await ScanRegistryKeyAsync(subKey, subKeyPath, hive, depth + 1);
                    }
                }
                catch (Exception ex)
                {
                    _errors.Add($"Error scanning subkey {subKeyName} in {keyPath}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            _errors.Add($"Error scanning key {keyPath}: {ex.Message}");
        }
    }

    /// <summary>
    /// Removes a specific registry entry.
    /// </summary>
    /// <param name="entry">The registry entry to remove.</param>
    /// <returns>True if removal was successful, false otherwise.</returns>
    private Task<bool> RemoveRegistryEntryAsync(RegistryEntry entry)
    {
        try
        {
            if (entry.EntryType == RegistryEntryType.Value)
            {
                return Task.FromResult(RemoveRegistryValue(entry));
            }

            return Task.FromResult(RemoveRegistryKey(entry));
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Exception removing registry entry {EntryPath}", entry.DisplayName);
            _errors.Add($"Error removing registry entry {entry.DisplayName}: {ex.Message}");
            return Task.FromResult(false);
        }
    }

    /// <summary>
    /// Removes a registry value from its containing key.
    /// </summary>
    /// <param name="entry">The registry value entry.</param>
    /// <returns>True if the value was removed or no longer exists.</returns>
    private bool RemoveRegistryValue(RegistryEntry entry)
    {
        if (entry.ValueName is null)
        {
            return false;
        }

        using var baseKey = OpenBaseKey(entry.RegistryHive);
        if (TryDeleteRegistryValue(baseKey, entry.KeyPath, entry.ValueName))
        {
            return true;
        }

        if (!TryGrantRegistryFullControl(baseKey, entry.KeyPath))
        {
            return false;
        }

        return TryDeleteRegistryValue(baseKey, entry.KeyPath, entry.ValueName);
    }

    /// <summary>
    /// Removes a registry key tree from its parent key.
    /// </summary>
    /// <param name="entry">The registry key entry.</param>
    /// <returns>True if the key was removed or no longer exists.</returns>
    private bool RemoveRegistryKey(RegistryEntry entry)
    {
        var parentPath = GetRegistryParentPath(entry.KeyPath);
        var keyName = GetRegistryLeafName(entry.KeyPath);

        if (string.IsNullOrWhiteSpace(parentPath) || string.IsNullOrWhiteSpace(keyName))
        {
            return false;
        }

        using var baseKey = OpenBaseKey(entry.RegistryHive);
        if (TryDeleteRegistryKey(baseKey, parentPath, keyName))
        {
            return true;
        }

        if (!TryGrantRegistryFullControl(baseKey, entry.KeyPath))
        {
            return false;
        }

        TryGrantRegistryFullControl(baseKey, parentPath);
        return TryDeleteRegistryKey(baseKey, parentPath, keyName);
    }

    /// <summary>
    /// Opens a registry base key for the specified hive.
    /// </summary>
    /// <param name="hive">The registry hive.</param>
    /// <returns>The opened base key.</returns>
    private static RegistryKey OpenBaseKey(RegistryHive hive)
    {
        return RegistryKey.OpenBaseKey(hive, RegistryView.Default);
    }

    /// <summary>
    /// Attempts to delete a registry value.
    /// </summary>
    /// <param name="baseKey">The hive base key.</param>
    /// <param name="keyPath">The containing key path.</param>
    /// <param name="valueName">The value to delete.</param>
    /// <returns>True if the value was removed or did not exist.</returns>
    private static bool TryDeleteRegistryValue(RegistryKey baseKey, string keyPath, string valueName)
    {
        try
        {
            using var key = baseKey.OpenSubKey(keyPath, writable: true);
            if (key is null)
            {
                return true;
            }

            if (key.GetValue(valueName) is null)
            {
                return true;
            }

            key.DeleteValue(valueName, throwOnMissingValue: false);
            return key.GetValue(valueName) is null;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Attempts to delete a registry key tree.
    /// </summary>
    /// <param name="baseKey">The hive base key.</param>
    /// <param name="parentPath">The parent key path.</param>
    /// <param name="keyName">The subkey to delete.</param>
    /// <returns>True if the key was removed or did not exist.</returns>
    private static bool TryDeleteRegistryKey(RegistryKey baseKey, string parentPath, string keyName)
    {
        try
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
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Attempts to take ownership of a registry key and grant the Administrators group full control.
    /// </summary>
    /// <param name="baseKey">The hive base key.</param>
    /// <param name="keyPath">The key path.</param>
    /// <returns>True if access was updated successfully.</returns>
    private bool TryGrantRegistryFullControl(RegistryKey baseKey, string keyPath)
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
        catch (Exception ex)
        {
            Logger.Debug(ex, "Failed to take ownership of registry key {KeyPath}: {Message}", keyPath, ex.Message);
            return false;
        }
    }

    /// <summary>
    /// Gets the parent registry path for a key.
    /// </summary>
    /// <param name="keyPath">The registry key path.</param>
    /// <returns>The parent registry path, or null when not available.</returns>
    private static string? GetRegistryParentPath(string keyPath)
    {
        var separatorIndex = keyPath.LastIndexOf('\\');
        return separatorIndex > 0 ? keyPath[..separatorIndex] : null;
    }

    /// <summary>
    /// Gets the leaf key name from a registry path.
    /// </summary>
    /// <param name="keyPath">The registry key path.</param>
    /// <returns>The final key name, or null when not available.</returns>
    private static string? GetRegistryLeafName(string keyPath)
    {
        var separatorIndex = keyPath.LastIndexOf('\\');
        return separatorIndex >= 0 && separatorIndex < keyPath.Length - 1
            ? keyPath[(separatorIndex + 1)..]
            : null;
    }

    /// <summary>
    /// Determines whether a registry path is the same as or nested below a parent path.
    /// </summary>
    /// <param name="candidatePath">The candidate registry path.</param>
    /// <param name="parentPath">The parent registry path.</param>
    /// <returns>True when the candidate is covered by the parent key removal.</returns>
    private static bool IsSameOrDescendantRegistryPath(string candidatePath, string parentPath)
    {
        if (candidatePath.Equals(parentPath, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return candidatePath.StartsWith(parentPath + "\\", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Stops Autodesk services that might interfere with registry operations.
    /// </summary>
    private async Task StopAutodeskServicesAsync()
    {
        var autodeskServices = new[] {
            "AdskLicensingService",
            "Autodesk Desktop App Service",
            "AdAppMgrSvc",
            "Mi-Service",
            "AdskAccessServiceHost"
        };

        Logger.Info("Stopping Autodesk services that might interfere with registry operations...");
        
        foreach (var serviceName in autodeskServices)
        {
            try
            {
                Logger.Debug("Attempting to stop service: {ServiceName}", serviceName);
                var startInfo = new ProcessStartInfo
                {
                    FileName = "sc.exe",
                    Arguments = $"stop \"{serviceName}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using var process = Process.Start(startInfo);
                if (process != null)
                {
                    await process.WaitForExitAsync();
                    if (process.ExitCode == 0)
                    {
                        Logger.Info("Successfully stopped service: {ServiceName}", serviceName);
                    }
                    else
                    {
                        Logger.Debug("Service {ServiceName} was not running or could not be stopped", serviceName);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Debug(ex, "Could not stop service {ServiceName}: {Message}", serviceName, ex.Message);
            }
        }
        
        // Wait a moment for services to fully stop
        await Task.Delay(2000);
        Logger.Info("Finished stopping Autodesk services");
    }

    /// <summary>
    /// Disposes of the scanner resources.
    /// </summary>
    public void Dispose()
    {
        if (!_disposed)
        {
            _foundEntries.Clear();
            _errors.Clear();
            _disposed = true;
        }
    }
}
