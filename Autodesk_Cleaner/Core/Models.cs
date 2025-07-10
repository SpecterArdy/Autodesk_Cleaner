using Microsoft.Win32;

namespace Autodesk_Cleaner.Core;

/// <summary>
/// Represents a Windows registry entry found during scanning.
/// </summary>
/// <param name="KeyPath">The full path to the registry key.</param>
/// <param name="ValueName">The name of the registry value (null for keys).</param>
/// <param name="ValueData">The data stored in the registry value.</param>
/// <param name="RegistryHive">The registry hive where the entry is located.</param>
/// <param name="EntryType">The type of registry entry.</param>
/// <param name="LastWriteTime">When the entry was last modified.</param>
public readonly record struct RegistryEntry(
    string KeyPath,
    string? ValueName,
    object? ValueData,
    RegistryHive RegistryHive,
    RegistryEntryType EntryType,
    DateTime LastWriteTime)
{
    /// <summary>
    /// Gets the display name for this registry entry.
    /// </summary>
    public readonly string DisplayName => 
        ValueName is not null ? $"{KeyPath}\\{ValueName}" : KeyPath;
    
    /// <summary>
    /// Gets the full registry path including hive.
    /// </summary>
    public readonly string FullPath => 
        $"{RegistryHive}\\{KeyPath}";
}

/// <summary>
/// Represents the type of registry entry.
/// </summary>
public enum RegistryEntryType
{
    /// <summary>
    /// A registry key.
    /// </summary>
    Key,
    
    /// <summary>
    /// A registry value.
    /// </summary>
    Value
}

/// <summary>
/// Represents the result of a registry removal operation.
/// </summary>
/// <param name="TotalEntries">The total number of entries processed.</param>
/// <param name="SuccessfulRemovals">The number of entries successfully removed.</param>
/// <param name="FailedRemovals">The number of entries that failed to be removed.</param>
/// <param name="Errors">Any errors encountered during removal.</param>
/// <param name="Duration">How long the operation took.</param>
public readonly record struct RemovalResult(
    int TotalEntries,
    int SuccessfulRemovals,
    int FailedRemovals,
    IReadOnlyCollection<string> Errors,
    TimeSpan Duration)
{
    /// <summary>
    /// Gets whether the operation was completely successful.
    /// </summary>
    public readonly bool IsSuccessful => FailedRemovals == 0;
    
    /// <summary>
    /// Gets the success rate as a percentage.
    /// </summary>
    public readonly double SuccessRate => TotalEntries == 0 ? 0.0 : (double)SuccessfulRemovals / TotalEntries * 100.0;
}

/// <summary>
/// Configuration options for the registry scanner.
/// </summary>
/// <param name="CreateBackup">Whether to create a backup before making changes.</param>
/// <param name="BackupPath">The path where backups should be saved.</param>
/// <param name="DryRun">Whether to perform a dry run without making actual changes.</param>
/// <param name="IncludeUserHive">Whether to include HKEY_CURRENT_USER in the scan.</param>
/// <param name="IncludeLocalMachine">Whether to include HKEY_LOCAL_MACHINE in the scan.</param>
/// <param name="MaxDepth">Maximum depth to scan in the registry tree.</param>
public readonly record struct ScannerConfig(
    bool CreateBackup = true,
    string BackupPath = "",
    bool DryRun = false,
    bool IncludeUserHive = true,
    bool IncludeLocalMachine = true,
    int MaxDepth = 10)
{
    /// <summary>
    /// Gets the default configuration for the scanner.
    /// </summary>
    public static ScannerConfig Default => new(
        CreateBackup: true,
        BackupPath: Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "RegistryBackup"),
        DryRun: false,
        IncludeUserHive: true,
        IncludeLocalMachine: true,
        MaxDepth: 10);
}
