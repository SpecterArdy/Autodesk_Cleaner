namespace Autodesk_Cleaner.Core;

/// <summary>
/// Interface for cleaning Autodesk-related files and directories from the file system.
/// </summary>
public interface IFileSystemCleaner
{
    /// <summary>
    /// Scans the file system for Autodesk-related files and directories.
    /// </summary>
    /// <returns>A collection of file system entries found.</returns>
    Task<IReadOnlyCollection<FileSystemEntry>> ScanFileSystemAsync();
    
    /// <summary>
    /// Removes the specified file system entries.
    /// </summary>
    /// <param name="entries">The file system entries to remove.</param>
    /// <returns>The result of the removal operation.</returns>
    Task<RemovalResult> RemoveEntriesAsync(IReadOnlyCollection<FileSystemEntry> entries);
    
    /// <summary>
    /// Creates a backup of important files before removal.
    /// </summary>
    /// <param name="backupPath">The path where the backup should be saved.</param>
    /// <param name="entries">The entries to backup.</param>
    /// <returns>True if backup was successful, false otherwise.</returns>
    Task<bool> CreateBackupAsync(string backupPath, IReadOnlyCollection<FileSystemEntry> entries);
    
    /// <summary>
    /// Validates file system entries before removal.
    /// </summary>
    /// <param name="entries">The entries to validate.</param>
    /// <returns>True if all entries are valid for removal, false otherwise.</returns>
    bool ValidateEntries(IReadOnlyCollection<FileSystemEntry> entries);
}

/// <summary>
/// Represents a file system entry found during scanning.
/// </summary>
/// <param name="Path">The full path to the file or directory.</param>
/// <param name="EntryType">The type of file system entry.</param>
/// <param name="Size">The size of the file or directory in bytes.</param>
/// <param name="LastModified">When the entry was last modified.</param>
/// <param name="IsHidden">Whether the entry is hidden.</param>
/// <param name="IsReadOnly">Whether the entry is read-only.</param>
public readonly record struct FileSystemEntry(
    string Path,
    FileSystemEntryType EntryType,
    long Size,
    DateTime LastModified,
    bool IsHidden,
    bool IsReadOnly)
{
    /// <summary>
    /// Gets the display name for this file system entry.
    /// </summary>
    public readonly string DisplayName => System.IO.Path.GetFileName(Path);
    
    /// <summary>
    /// Gets the parent directory path.
    /// </summary>
    public readonly string? ParentDirectory => System.IO.Path.GetDirectoryName(Path);
}

/// <summary>
/// Represents the type of file system entry.
/// </summary>
public enum FileSystemEntryType
{
    /// <summary>
    /// A file.
    /// </summary>
    File,
    
    /// <summary>
    /// A directory.
    /// </summary>
    Directory
}
