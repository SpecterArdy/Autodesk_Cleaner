using Microsoft.Win32;

namespace Autodesk_Cleaner.Core;

/// <summary>
/// Interface for scanning and managing Windows registry entries related to Autodesk products.
/// </summary>
public interface IRegistryScanner
{
    /// <summary>
    /// Scans the registry for Autodesk-related entries.
    /// </summary>
    /// <returns>A collection of registry entries found.</returns>
    Task<IReadOnlyCollection<RegistryEntry>> ScanRegistryAsync();
    
    /// <summary>
    /// Removes the specified registry entries.
    /// </summary>
    /// <param name="entries">The registry entries to remove.</param>
    /// <returns>The result of the removal operation.</returns>
    Task<RemovalResult> RemoveEntriesAsync(IReadOnlyCollection<RegistryEntry> entries);
    
    /// <summary>
    /// Creates a backup of the registry before making changes.
    /// </summary>
    /// <param name="backupPath">The path where the backup should be saved.</param>
    /// <returns>True if backup was successful, false otherwise.</returns>
    Task<bool> CreateBackupAsync(string backupPath);
    
    /// <summary>
    /// Validates registry entries before removal.
    /// </summary>
    /// <param name="entries">The entries to validate.</param>
    /// <returns>True if all entries are valid for removal, false otherwise.</returns>
    bool ValidateEntries(IReadOnlyCollection<RegistryEntry> entries);
}
