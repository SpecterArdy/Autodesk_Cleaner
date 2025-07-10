using System.Diagnostics;
using System.IO.Compression;
using System.IO.Abstractions;
using System.Text.RegularExpressions;
using NLog;
using Spectre.Console;
namespace Autodesk_Cleaner.Core;

/// <summary>
/// Implementation of IFileSystemCleaner specifically designed for Autodesk file cleanup.
/// </summary>
public sealed class AutodeskFileSystemCleaner : IFileSystemCleaner, IDisposable
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
    private readonly ScannerConfig _config;
    private readonly List<FileSystemEntry> _foundEntries = [];
    private readonly List<string> _errors = [];
    private bool _disposed;

    /// <summary>
    /// Autodesk file system paths to scan and clean.
    /// </summary>
    private static readonly IReadOnlyList<string> AutodeskPaths = [
        @"C:\Program Files\Autodesk",
        @"C:\Program Files\Common Files\Autodesk Shared",
        @"C:\Program Files (x86)\Autodesk",
        @"C:\Program Files (x86)\Common Files\Autodesk Shared",
        @"C:\ProgramData\Autodesk",
        @"C:\ProgramData\FLEXnet"
    ];

    /// <summary>
    /// User-specific Autodesk paths to scan and clean.
    /// </summary>
    private static readonly IReadOnlyList<string> UserAutodeskPaths = [
        @"Autodesk",
        @"AppData\Local\Autodesk",
        @"AppData\Roaming\Autodesk"
    ];

    /// <summary>
    /// Temp directory paths to clean.
    /// </summary>
    private static readonly IReadOnlyList<string> TempPaths = [
        @"C:\Windows\Temp",
        @"C:\Users\{username}\AppData\Local\Temp"
    ];

    /// <summary>
    /// Patterns to match Autodesk-related files and directories.
    /// </summary>
    private static readonly Regex AutodeskPattern = new(
        @"(autodesk|maya|3ds\s*max|autocad|revit|inventor|fusion|vault|navisworks|mudbox|motionbuilder|alias|adsk)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Patterns to match ADSK files in FLEXnet directory.
    /// </summary>
    private static readonly Regex FlexNetPattern = new(
        @"^adsk.*",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Initializes a new instance of the AutodeskFileSystemCleaner.
    /// </summary>
    /// <param name="config">Configuration options for the cleaner.</param>
    public AutodeskFileSystemCleaner(ScannerConfig config = default)
    {
        _config = config == default ? ScannerConfig.Default : config;
        Logger.Info("AutodeskFileSystemCleaner initialized with config: {@Config}", _config);
    }

    /// <inheritdoc />
    public async Task<IReadOnlyCollection<FileSystemEntry>> ScanFileSystemAsync()
    {
        Logger.Info("Starting file system scan with configuration: {@Config}", _config);
        _foundEntries.Clear();
        _errors.Clear();

        try
        {
            // Scan system-wide Autodesk paths
            Logger.Info("Scanning system-wide Autodesk paths");
            foreach (var path in AutodeskPaths)
            {
                Logger.Debug("Scanning system path: {Path}", path);
                await ScanDirectoryAsync(path);
            }
            Logger.Info("Completed system-wide scan, found {EntryCount} entries", _foundEntries.Count);

            // Scan user-specific Autodesk paths
            Logger.Info("Scanning user-specific Autodesk paths");
            var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            var initialCount = _foundEntries.Count;
            foreach (var relativePath in UserAutodeskPaths)
            {
                var fullPath = Path.Combine(userProfile, relativePath);
                Logger.Debug("Scanning user path: {Path}", fullPath);
                await ScanDirectoryAsync(fullPath);
            }
            Logger.Info("Completed user-specific scan, found {NewEntryCount} additional entries", _foundEntries.Count - initialCount);

            // Scan temp directories for Autodesk files
            Logger.Info("Scanning temp directories for Autodesk files");
            initialCount = _foundEntries.Count;
            await ScanTempDirectoriesAsync();
            Logger.Info("Completed temp directory scan, found {NewEntryCount} additional entries", _foundEntries.Count - initialCount);

            // Scan FLEXnet directory for ADSK files
            Logger.Info("Scanning FLEXnet directory for ADSK files");
            initialCount = _foundEntries.Count;
            await ScanFlexNetDirectoryAsync();
            Logger.Info("Completed FLEXnet scan, found {NewEntryCount} additional entries", _foundEntries.Count - initialCount);

            Logger.Info("File system scan completed. Total entries found: {TotalEntries}, Errors: {ErrorCount}", _foundEntries.Count, _errors.Count);
            return _foundEntries.AsReadOnly();
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Critical error during file system scan");
            _errors.Add($"Critical error during file system scan: {ex.Message}");
            throw new InvalidOperationException("File system scan failed", ex);
        }
    }

    /// <inheritdoc />
    public async Task<RemovalResult> RemoveEntriesAsync(IReadOnlyCollection<FileSystemEntry> entries)
    {
        // Filter to only valid entries
        var validEntries = FilterValidEntries(entries);
        
        if (validEntries.Count == 0)
        {
            Logger.Error("File system entry validation failed - no valid entries found to process");
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
            var backupPath = Path.Combine(_config.BackupPath, $"files_backup_{DateTime.Now:yyyyMMdd_HHmmss}.zip");
            if (!await CreateBackupAsync(backupPath, entries))
            {
                errors.Add("Failed to create file system backup");
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
        }

        // Process each valid entry
        if (!validEntries.Any())
        {
            Logger.Warn("No valid file system entries to process.");
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
                    AnsiConsole.MarkupLine($"[dim][DRY RUN] Would remove: {entry.Path}[/]");
                    successfulRemovals++;
                    continue;
                }

                if (await RemoveFileSystemEntryAsync(entry))
                {
                    successfulRemovals++;
                }
                else
                {
                    errors.Add($"Failed to remove: {entry.Path}");
                }
            }
            catch (Exception ex)
            {
                errors.Add($"Error removing {entry.Path}: {ex.Message}");
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
    public async Task<bool> CreateBackupAsync(string backupPath, IReadOnlyCollection<FileSystemEntry> entries)
    {
        try
        {
            var directory = Path.GetDirectoryName(backupPath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using var zipFile = new FileStream(backupPath, FileMode.Create);
            using var archive = new ZipArchive(zipFile, ZipArchiveMode.Create);

            foreach (var entry in entries.Where(e => e.EntryType == FileSystemEntryType.File))
            {
                try
                {
                    if (File.Exists(entry.Path))
                    {
                        var relativePath = entry.Path.Replace(@":\", "_").Replace(@"\", "_");
                        var zipEntry = archive.CreateEntry(relativePath);
                        
                        using var fileStream = new FileStream(entry.Path, FileMode.Open, FileAccess.Read);
                        using var zipStream = zipEntry.Open();
                        await fileStream.CopyToAsync(zipStream);
                    }
                }
                catch (Exception ex)
                {
                    _errors.Add($"Error backing up {entry.Path}: {ex.Message}");
                }
            }

            return true;
        }
        catch (Exception ex)
        {
            _errors.Add($"Backup creation failed: {ex.Message}");
            return false;
        }
    }

    /// <inheritdoc />
    public bool ValidateEntries(IReadOnlyCollection<FileSystemEntry> entries)
    {
        if (entries is null || entries.Count == 0)
        {
            Logger.Warn("ValidateEntries failed: entries is null or empty");
            return false;
        }

        Logger.Info("Validating {EntryCount} file system entries", entries.Count);
        
        // Check each entry and count valid/invalid for debugging
        var validEntries = new List<FileSystemEntry>();
        var invalidEntries = new List<FileSystemEntry>();
        
        foreach (var entry in entries)
        {
            var isAutodeskPath = IsAutodeskPath(entry.Path);
            var matchesPath = AutodeskPattern.IsMatch(entry.Path);
            var matchesFileName = AutodeskPattern.IsMatch(Path.GetFileName(entry.Path));
            var matchesFlexNet = FlexNetPattern.IsMatch(Path.GetFileName(entry.Path));
            
            var isValid = isAutodeskPath || matchesPath || matchesFileName || matchesFlexNet;
            
            if (isValid)
            {
                validEntries.Add(entry);
            }
            else
            {
                invalidEntries.Add(entry);
                Logger.Debug("Invalid entry: {Path}, AutodeskPath: {IsAutodeskPath}, PathMatch: {MatchesPath}, FileNameMatch: {MatchesFileName}, FlexNetMatch: {MatchesFlexNet}", 
                    entry.Path, isAutodeskPath, matchesPath, matchesFileName, matchesFlexNet);
            }
        }
        
        Logger.Info("Validation results: {ValidCount} valid entries, {InvalidCount} invalid entries out of {TotalCount} total", 
            validEntries.Count, invalidEntries.Count, entries.Count);
        
        if (invalidEntries.Count > 0)
        {
            Logger.Warn("Found {InvalidCount} invalid entries out of {TotalCount} - these will be filtered out", invalidEntries.Count, entries.Count);
            // Log first few invalid entries for debugging
            foreach (var entry in invalidEntries.Take(3))
            {
                Logger.Debug("Invalid entry example: {Path}", entry.Path);
            }
        }
        
        // Return true if we have ANY valid entries to process
        var hasValidEntries = validEntries.Count > 0;
        Logger.Info("Validation result: {HasValidEntries} - {ValidCount} entries will be processed", 
            hasValidEntries ? "PASS" : "FAIL", validEntries.Count);
        
        return hasValidEntries;
    }

    /// <summary>
    /// Gets the count of valid Autodesk-related entries from a collection.
    /// </summary>
    /// <param name="entries">The file system entries to analyze.</param>
    /// <returns>The number of valid entries.</returns>
    public int GetValidEntryCount(IReadOnlyCollection<FileSystemEntry> entries)
    {
        Logger.Info("GetValidEntryCount: Starting count for {EntryCount} file system entries", entries?.Count ?? 0);
        if (entries == null || entries.Count == 0)
        {
            Logger.Warn("GetValidEntryCount: Null or empty entries provided");
            return 0;
        }
        var validEntries = FilterValidEntries(entries);
        Logger.Info("GetValidEntryCount: Found {ValidCount} valid file system entries", validEntries.Count);
        return validEntries.Count;
    }

    /// <summary>
    /// Filters file system entries to only include valid Autodesk-related entries.
    /// </summary>
    /// <param name="entries">The file system entries to filter.</param>
    /// <returns>A list of valid Autodesk-related entries.</returns>
    private List<FileSystemEntry> FilterValidEntries(IReadOnlyCollection<FileSystemEntry> entries)
    {
        if (entries is null || entries.Count == 0)
        {
            Logger.Info("FilterValidEntries: entries is null or empty");
            return new List<FileSystemEntry>();
        }

        Logger.Info("FilterValidEntries: Processing {EntryCount} file system entries", entries.Count);
        var validEntries = new List<FileSystemEntry>();
        var invalidCount = 0;
        
        foreach (var entry in entries)
        {
            var isAutodeskPath = IsAutodeskPath(entry.Path);
            var matchesPath = AutodeskPattern.IsMatch(entry.Path);
            var matchesFileName = AutodeskPattern.IsMatch(Path.GetFileName(entry.Path));
            var matchesFlexNet = FlexNetPattern.IsMatch(Path.GetFileName(entry.Path));
            
            var isValid = isAutodeskPath || matchesPath || matchesFileName || matchesFlexNet;
            
            if (isValid)
            {
                validEntries.Add(entry);
            }
            else
            {
                invalidCount++;
                if (invalidCount <= 3) // Log first 3 invalid entries for debugging
                {
                    Logger.Debug("FilterValidEntries: Invalid file entry {Path} - AutodeskPath: {IsAutodeskPath}, PathMatch: {MatchesPath}, FileNameMatch: {MatchesFileName}, FlexNetMatch: {MatchesFlexNet}", 
                        entry.Path, isAutodeskPath, matchesPath, matchesFileName, matchesFlexNet);
                }
            }
        }
        
        Logger.Info("FilterValidEntries: Found {ValidCount} valid file system entries, {InvalidCount} invalid entries", validEntries.Count, invalidCount);
        return validEntries;
    }

    /// <summary>
    /// Checks if a file system path is under a known Autodesk directory.
    /// </summary>
    /// <param name="path">The file system path to check.</param>
    /// <returns>True if the path is under an Autodesk directory.</returns>
    private static bool IsAutodeskPath(string path)
    {
        if (string.IsNullOrEmpty(path))
            return false;

        // Check if path is under any of the known Autodesk paths
        if (AutodeskPaths.Any(basePath => 
            path.StartsWith(basePath, StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        // Check if path is under any user-specific Autodesk paths
        var userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        return UserAutodeskPaths.Any(relativePath => 
        {
            var fullPath = Path.Combine(userProfile, relativePath);
            return path.StartsWith(fullPath, StringComparison.OrdinalIgnoreCase);
        });
    }

    /// <summary>
    /// Scans a directory for Autodesk-related files and directories.
    /// </summary>
    /// <param name="directoryPath">The directory path to scan.</param>
    private async Task ScanDirectoryAsync(string directoryPath)
    {
        try
        {
            if (!Directory.Exists(directoryPath))
            {
                return;
            }

            var directoryInfo = new DirectoryInfo(directoryPath);
            var attributes = directoryInfo.Attributes;

            // Add the directory itself if it's Autodesk-related
            if (AutodeskPattern.IsMatch(directoryPath))
            {
                _foundEntries.Add(new FileSystemEntry(
                    Path: directoryPath,
                    EntryType: FileSystemEntryType.Directory,
                    Size: await GetDirectorySizeAsync(directoryPath),
                    LastModified: directoryInfo.LastWriteTime,
                    IsHidden: attributes.HasFlag(FileAttributes.Hidden),
                    IsReadOnly: attributes.HasFlag(FileAttributes.ReadOnly)));
            }

            // Scan files in the directory
            foreach (var file in directoryInfo.GetFiles("*", SearchOption.AllDirectories))
            {
                try
                {
                    if (AutodeskPattern.IsMatch(file.FullName) || 
                        AutodeskPattern.IsMatch(file.Name))
                    {
                        _foundEntries.Add(new FileSystemEntry(
                            Path: file.FullName,
                            EntryType: FileSystemEntryType.File,
                            Size: file.Length,
                            LastModified: file.LastWriteTime,
                            IsHidden: file.Attributes.HasFlag(FileAttributes.Hidden),
                            IsReadOnly: file.Attributes.HasFlag(FileAttributes.ReadOnly)));
                    }
                }
                catch (Exception ex)
                {
                    _errors.Add($"Error scanning file {file.FullName}: {ex.Message}");
                }
            }

            // Scan subdirectories
            foreach (var subDirectory in directoryInfo.GetDirectories("*", SearchOption.AllDirectories))
            {
                try
                {
                    if (AutodeskPattern.IsMatch(subDirectory.FullName) || 
                        AutodeskPattern.IsMatch(subDirectory.Name))
                    {
                        _foundEntries.Add(new FileSystemEntry(
                            Path: subDirectory.FullName,
                            EntryType: FileSystemEntryType.Directory,
                            Size: await GetDirectorySizeAsync(subDirectory.FullName),
                            LastModified: subDirectory.LastWriteTime,
                            IsHidden: subDirectory.Attributes.HasFlag(FileAttributes.Hidden),
                            IsReadOnly: subDirectory.Attributes.HasFlag(FileAttributes.ReadOnly)));
                    }
                }
                catch (Exception ex)
                {
                    _errors.Add($"Error scanning directory {subDirectory.FullName}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            _errors.Add($"Error scanning directory {directoryPath}: {ex.Message}");
        }
    }

    /// <summary>
    /// Scans temp directories for Autodesk-related files.
    /// </summary>
    private Task ScanTempDirectoriesAsync()
    {
        var tempPaths = new List<string>
        {
            Path.GetTempPath(),
            Environment.GetEnvironmentVariable("TEMP") ?? string.Empty,
            Environment.GetEnvironmentVariable("TMP") ?? string.Empty
        };

        foreach (var tempPath in tempPaths.Where(p => !string.IsNullOrEmpty(p) && Directory.Exists(p)))
        {
            try
            {
                var tempDir = new DirectoryInfo(tempPath);
                foreach (var item in tempDir.GetFileSystemInfos("*", SearchOption.TopDirectoryOnly))
                {
                    if (AutodeskPattern.IsMatch(item.Name))
                    {
                        if (item is FileInfo file)
                        {
                            _foundEntries.Add(new FileSystemEntry(
                                Path: file.FullName,
                                EntryType: FileSystemEntryType.File,
                                Size: file.Length,
                                LastModified: file.LastWriteTime,
                                IsHidden: file.Attributes.HasFlag(FileAttributes.Hidden),
                                IsReadOnly: file.Attributes.HasFlag(FileAttributes.ReadOnly)));
                        }
                        else if (item is DirectoryInfo directory)
                        {
                            _foundEntries.Add(new FileSystemEntry(
                                Path: directory.FullName,
                                EntryType: FileSystemEntryType.Directory,
                        Size: GetDirectorySizeAsync(directory.FullName).Result,
                                LastModified: directory.LastWriteTime,
                                IsHidden: directory.Attributes.HasFlag(FileAttributes.Hidden),
                                IsReadOnly: directory.Attributes.HasFlag(FileAttributes.ReadOnly)));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _errors.Add($"Error scanning temp directory {tempPath}: {ex.Message}");
            }
        }
        
        return Task.CompletedTask;
    }

    /// <summary>
    /// Scans the FLEXnet directory for ADSK files.
    /// </summary>
    private Task ScanFlexNetDirectoryAsync()
    {
        var flexNetPath = @"C:\ProgramData\FLEXnet";
        
        try
        {
            if (!Directory.Exists(flexNetPath))
            {
                return Task.CompletedTask;
            }

            var flexNetDir = new DirectoryInfo(flexNetPath);
            foreach (var file in flexNetDir.GetFiles("*", SearchOption.AllDirectories))
            {
                if (FlexNetPattern.IsMatch(file.Name))
                {
                    _foundEntries.Add(new FileSystemEntry(
                        Path: file.FullName,
                        EntryType: FileSystemEntryType.File,
                        Size: file.Length,
                        LastModified: file.LastWriteTime,
                        IsHidden: file.Attributes.HasFlag(FileAttributes.Hidden),
                        IsReadOnly: file.Attributes.HasFlag(FileAttributes.ReadOnly)));
                }
            }
        }
        catch (Exception ex)
        {
            _errors.Add($"Error scanning FLEXnet directory: {ex.Message}");
        }
        
        return Task.CompletedTask;
    }

    /// <summary>
    /// Calculates the total size of a directory.
    /// </summary>
    /// <param name="directoryPath">The directory path.</param>
    /// <returns>The total size in bytes.</returns>
    private Task<long> GetDirectorySizeAsync(string directoryPath)
    {
        try
        {
            if (!Directory.Exists(directoryPath))
            {
                return Task.FromResult(0L);
            }

            var directoryInfo = new DirectoryInfo(directoryPath);
            var size = directoryInfo.GetFiles("*", SearchOption.AllDirectories).Sum(file => file.Length);
            return Task.FromResult(size);
        }
        catch
        {
            return Task.FromResult(0L);
        }
    }

    /// <summary>
    /// Removes a specific file system entry.
    /// </summary>
    /// <param name="entry">The file system entry to remove.</param>
    /// <returns>True if removal was successful, false otherwise.</returns>
    private Task<bool> RemoveFileSystemEntryAsync(FileSystemEntry entry)
    {
        try
        {
            if (entry.EntryType == FileSystemEntryType.File)
            {
                if (File.Exists(entry.Path))
                {
                    // Remove read-only attribute if present
                    if (entry.IsReadOnly)
                    {
                        File.SetAttributes(entry.Path, File.GetAttributes(entry.Path) & ~FileAttributes.ReadOnly);
                    }

                    File.Delete(entry.Path);
                    return Task.FromResult(true);
                }
            }
            else if (entry.EntryType == FileSystemEntryType.Directory)
            {
                if (Directory.Exists(entry.Path))
                {
                    // Remove read-only attributes from all files in the directory
                    var directoryInfo = new DirectoryInfo(entry.Path);
                    foreach (var file in directoryInfo.GetFiles("*", SearchOption.AllDirectories))
                    {
                        if (file.IsReadOnly)
                        {
                            file.Attributes &= ~FileAttributes.ReadOnly;
                        }
                    }

                    Directory.Delete(entry.Path, true);
                    return Task.FromResult(true);
                }
            }

            return Task.FromResult(false);
        }
        catch (Exception ex)
        {
            _errors.Add($"Error removing file system entry {entry.Path}: {ex.Message}");
            return Task.FromResult(false);
        }
    }

    /// <summary>
    /// Disposes of the cleaner resources.
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
