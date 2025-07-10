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
            return _foundEntries.AsReadOnly();
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
        if (!ValidateEntries(entries))
        {
            throw new ArgumentException("Invalid entries provided for removal", nameof(entries));
        }

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

        // Process each entry
        foreach (var entry in entries)
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
            TotalEntries: entries.Count,
            SuccessfulRemovals: successfulRemovals,
            FailedRemovals: entries.Count - successfulRemovals,
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
            return false;
        }

        // Validate that all entries are Autodesk-related
        return entries.All(entry => 
            AutodeskPattern.IsMatch(entry.KeyPath) || 
            (entry.ValueName is not null && AutodeskPattern.IsMatch(entry.ValueName)) ||
            (entry.ValueData is string valueData && AutodeskPattern.IsMatch(valueData)));
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
