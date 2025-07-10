using Microsoft.Win32;
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
        @"SOFTWARE\Classes\Installer\Products"
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
        
        if (validEntries.Count == 0)
        {
            Logger.Error("Registry entry validation failed - no valid entries found to process");
            return new RemovalResult(
                TotalEntries: 0,
                SuccessfulRemovals: 0,
                FailedRemovals: 0,
                Errors: ["Entry validation failed - no Autodesk-related entries found to process"],
                Duration: TimeSpan.Zero);
        }
        
        Logger.Info("Processing {ValidCount} valid entries out of {TotalCount} total entries", validEntries.Count, entries.Count);

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
                Logger.Error("Failed to create registry backup at: {BackupPath}", backupPath);
                errors.Add("Failed to create registry backup");
                if (!_config.DryRun)
                {
                    return new RemovalResult(
                        TotalEntries: entries.Count,
                        SuccessfulRemovals: 0,
                        FailedRemovals: entries.Count,
                        Errors: errors,
                        Duration: stopwatch.Elapsed);
                }
            }
            else
            {
                Logger.Info("Registry backup created successfully at: {BackupPath}", backupPath);
            }
        }

        // Process each valid entry
        if (!validEntries.Any())
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

        foreach (var entry in validEntries)
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
                errors.Add($"Error removing {entry.DisplayName}: {ex.Message}");
            }
        }

        stopwatch.Stop();

        return new RemovalResult(
            TotalEntries: validEntries.Count,
            SuccessfulRemovals: successfulRemovals,
            FailedRemovals: validEntries.Count - successfulRemovals,
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
                Directory.CreateDirectory(directory);
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
            var hiveKey = entry.RegistryHive switch
            {
                RegistryHive.LocalMachine => Registry.LocalMachine,
                RegistryHive.CurrentUser => Registry.CurrentUser,
                _ => throw new NotSupportedException($"Registry hive {entry.RegistryHive} not supported")
            };

            if (entry.EntryType == RegistryEntryType.Value)
            {
                // Remove registry value
                var keyPath = Path.GetDirectoryName(entry.KeyPath.Replace('\\', Path.DirectorySeparatorChar))?.Replace(Path.DirectorySeparatorChar, '\\');
                if (keyPath is not null)
                {
                    using var key = hiveKey.OpenSubKey(keyPath, true);
                    if (key is not null && entry.ValueName is not null)
                    {
                        key.DeleteValue(entry.ValueName, false);
                        return Task.FromResult(true);
                    }
                }
            }
            else
            {
                // Remove registry key
                var parentPath = Path.GetDirectoryName(entry.KeyPath.Replace('\\', Path.DirectorySeparatorChar))?.Replace(Path.DirectorySeparatorChar, '\\');
                var keyName = Path.GetFileName(entry.KeyPath.Replace('\\', Path.DirectorySeparatorChar));
                
                if (parentPath is not null && keyName is not null)
                {
                    using var parentKey = hiveKey.OpenSubKey(parentPath, true);
                    if (parentKey is not null)
                    {
                        parentKey.DeleteSubKeyTree(keyName, false);
                        return Task.FromResult(true);
                    }
                }
            }

            return Task.FromResult(false);
        }
        catch (Exception ex)
        {
            _errors.Add($"Error removing registry entry {entry.DisplayName}: {ex.Message}");
            return Task.FromResult(false);
        }
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
