using System.Diagnostics;
using System.IO.Compression;
using System.IO.Abstractions;
using System.Management;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text.RegularExpressions;
using Microsoft.Management.Infrastructure;
using NLog;
using Spectre.Console;
namespace Autodesk_Cleaner.Core;

/// <summary>
/// Implementation of IFileSystemCleaner specifically designed for Autodesk file cleanup.
/// </summary>
public sealed class AutodeskFileSystemCleaner : IFileSystemCleaner, IDisposable
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
    
    // Windows API declarations for file handle detection
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr OpenProcess(uint processAccess, bool bInheritHandle, int processId);
    
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr hObject);
    
    private const uint PROCESS_QUERY_INFORMATION = 0x0400;
    private const uint PROCESS_VM_READ = 0x0010;
    
    /// <summary>
    /// Critical Windows system processes that should never be killed.
    /// </summary>
    private static readonly HashSet<string> ProtectedProcesses = new(StringComparer.OrdinalIgnoreCase)
    {
        "System", "Registry", "smss", "csrss", "wininit", "winlogon", "services", "lsass", "lsm", "svchost",
        "explorer", "dwm", "conhost", "audiodg", "spoolsv", "SearchIndexer", "WmiPrvSE", "dllhost",
        "RuntimeBroker", "ShellExperienceHost", "SearchUI", "StartMenuExperienceHost", "Taskmgr",
        "MsMpEng", "NisSrv", "SecurityHealthService", "windefend", "MpDefenderCoreService"
    };
    private readonly ScannerConfig _config;
    private readonly List<FileSystemEntry> _foundEntries = [];
    private readonly List<string> _errors = [];
    private bool _disposed;

    /// <summary>
    /// Autodesk file system paths to scan and clean.
    /// </summary>
    private static readonly IReadOnlyList<string> AutodeskPaths = [
        @"C:\Autodesk",
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
    /// Known Autodesk installer and deployment cache paths, including 2027 media extracted by modern installers.
    /// </summary>
    private static readonly IReadOnlyList<string> AutodeskInstallMediaPaths = [
        @"C:\Autodesk\WI",
        @"C:\ProgramData\Autodesk\ODIS",
        @"C:\Program Files\Autodesk\AdODIS",
        @"C:\Program Files (x86)\Autodesk\AdODIS"
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
        @"(autodesk|maya|3ds\s*max|autocad|revit|inventor|fusion|vault|navisworks|mudbox|motionbuilder|alias|adsk|autodesk\s*access|adskaccess|odis|adodis)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Patterns that identify Autodesk 2027 installer artifacts even when the file name is generic.
    /// </summary>
    private static readonly Regex Autodesk2027InstallPattern = new(
        @"(?<!\d)(2027|R27)(?!\d)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Patterns to match ADSK files in FLEXnet directory.
    /// </summary>
    private static readonly Regex FlexNetPattern = new(
        @"^adsk.*",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Enumeration settings used when clearing attributes or force-deleting directory trees.
    /// </summary>
    private static readonly System.IO.EnumerationOptions RecursiveEnumerationOptions = new()
    {
        RecurseSubdirectories = true,
        IgnoreInaccessible = true,
        AttributesToSkip = 0,
        ReturnSpecialDirectories = false
    };

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
            // Return a copy to avoid issues with disposal clearing the internal list
            return _foundEntries.ToList().AsReadOnly();
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
        var removalPlan = BuildRemovalPlan(validEntries);
        
        if (removalPlan.Count == 0)
        {
            Logger.Error("File system entry validation failed - no valid entries found to process");
            return new RemovalResult(
                TotalEntries: 0,
                SuccessfulRemovals: 0,
                FailedRemovals: 0,
                Errors: ["Entry validation failed - no Autodesk-related entries found to process"],
                Duration: TimeSpan.Zero);
        }
        
        Logger.Info("Processing {RemovalPlanCount} planned file system removals from {ValidCount} valid entries out of {TotalCount} total entries",
            removalPlan.Count, validEntries.Count, entries.Count);

        var stopwatch = Stopwatch.StartNew();
        var errors = new List<string>();
        var successfulRemovals = 0;

        // Stop Autodesk services before deletion
        if (!_config.DryRun)
        {
            AnsiConsole.MarkupLine("[yellow]Stopping Autodesk services that can block deletion...[/]");
        }
        await StopAutodeskServicesAsync();
        
        // Kill any remaining Autodesk processes
        if (!_config.DryRun)
        {
            AnsiConsole.MarkupLine("[yellow]Killing Autodesk processes still holding files open...[/]");
        }
        await KillAllAutodeskProcessesAsync();

        if (!_config.DryRun)
        {
            AnsiConsole.MarkupLine($"[yellow]Beginning removal plan for {removalPlan.Count} top-level entries...[/]");
        }
        
        // Process each valid entry
        if (!removalPlan.Any())
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

        // Progress reporting setup
        var processedCount = 0;
        var totalCount = removalPlan.Count;
        var completedDirectoryRemovals = 0;
        var completedFileRemovals = 0;
        var failedRemovals = 0;
        
        Logger.Info("Starting removal of {TotalCount} file system entries", totalCount);

        if (!_config.DryRun)
        {
            AnsiConsole.MarkupLine(
                $"[bold yellow]Cleanup progress:[/] 0/{totalCount} processed | directories deleted: 0 | files deleted: 0 | failed: 0 | ETA: calculating...");
        }
        
        foreach (var entry in removalPlan)
        {
            try
            {
                processedCount++;
                var operationLabel = entry.EntryType == FileSystemEntryType.Directory ? "directory tree" : "file";
                var progressPercentage = (double)processedCount / totalCount * 100;
                var etaText = FormatEta(stopwatch.Elapsed, processedCount, totalCount);
                
                if (!_config.DryRun)
                {
                    Logger.Info("Processing file system entries: {ProcessedCount}/{TotalCount} ({ProgressPercentage:F1}%) ETA {Eta}",
                        processedCount, totalCount, progressPercentage, etaText);
                    AnsiConsole.MarkupLine(
                        $"[cyan]Working on {operationLabel} {processedCount}/{totalCount} ({progressPercentage:F1}%):[/] {Markup.Escape(entry.Path)} [dim]| ETA {etaText}[/]");
                }
                
                if (_config.DryRun)
                {
                    AnsiConsole.MarkupLine($"[dim][DRY RUN] Would remove: {entry.Path}[/]");
                    successfulRemovals++;
                    
                    // Yield control periodically during dry run
                    if (processedCount % 100 == 0)
                    {
                        await Task.Delay(1); // Yield control to prevent UI freezing
                    }
                    continue;
                }

                Logger.Debug("Removing file system entry: {Path}", entry.Path);
                if (await RemoveFileSystemEntryAsync(entry))
                {
                    Logger.Debug("Successfully removed file system entry: {Path}", entry.Path);
                    successfulRemovals++;
                    if (entry.EntryType == FileSystemEntryType.Directory)
                    {
                        completedDirectoryRemovals++;
                    }
                    else
                    {
                        completedFileRemovals++;
                    }

                    if (!_config.DryRun)
                    {
                        AnsiConsole.MarkupLine(
                            $"[green]Completed {operationLabel} removal {processedCount}/{totalCount}:[/] {Markup.Escape(entry.Path)}");
                        AnsiConsole.MarkupLine(
                            $"[dim]Progress:[/] {processedCount}/{totalCount} processed | directories deleted: {completedDirectoryRemovals} | files deleted: {completedFileRemovals} | failed: {failedRemovals} | ETA: {FormatEta(stopwatch.Elapsed, processedCount, totalCount)}");
                    }
                }
                else
                {
                    Logger.Warn("Failed to remove file system entry: {Path}", entry.Path);
                    errors.Add($"Failed to remove: {entry.Path}");
                    failedRemovals++;

                    if (!_config.DryRun)
                    {
                        AnsiConsole.MarkupLine(
                            $"[red]Failed {operationLabel} removal {processedCount}/{totalCount}:[/] {Markup.Escape(entry.Path)}");
                        AnsiConsole.MarkupLine(
                            $"[dim]Progress:[/] {processedCount}/{totalCount} processed | directories deleted: {completedDirectoryRemovals} | files deleted: {completedFileRemovals} | failed: {failedRemovals} | ETA: {FormatEta(stopwatch.Elapsed, processedCount, totalCount)}");
                    }
                }
                
                // Yield control periodically to prevent UI freezing
                if (processedCount % 10 == 0)
                {
                    await Task.Delay(1); // Small delay to allow UI updates
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Error removing file system entry {Path}: {Message}", entry.Path, ex.Message);
                errors.Add($"Error removing {entry.Path}: {ex.Message}");
                failedRemovals++;

                if (!_config.DryRun)
                {
                    AnsiConsole.MarkupLine(
                        $"[red]Error removing entry {processedCount}/{totalCount}:[/] {Markup.Escape(entry.Path)} [dim]({Markup.Escape(ex.Message)})[/]");
                    AnsiConsole.MarkupLine(
                        $"[dim]Progress:[/] {processedCount}/{totalCount} processed | directories deleted: {completedDirectoryRemovals} | files deleted: {completedFileRemovals} | failed: {failedRemovals} | ETA: {FormatEta(stopwatch.Elapsed, processedCount, totalCount)}");
                }
            }
        }
        
        Logger.Info("Completed processing {TotalCount} file system entries. Successful: {SuccessfulCount}, Failed: {FailedCount}", 
            totalCount, successfulRemovals, totalCount - successfulRemovals);

        stopwatch.Stop();

        return new RemovalResult(
            TotalEntries: removalPlan.Count,
            SuccessfulRemovals: successfulRemovals,
            FailedRemovals: removalPlan.Count - successfulRemovals,
            Errors: errors,
            Duration: stopwatch.Elapsed);
    }

    /// <inheritdoc />
    public async Task<bool> CreateBackupAsync(string backupPath, IReadOnlyCollection<FileSystemEntry> entries)
    {
        Logger.Info("File system backup creation is disabled; skipping backup request for {EntryCount} entries", entries.Count);
        await Task.CompletedTask;
        return true;
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
            var isValid = MatchesAutodeskTarget(entry.Path);
            
            if (isValid)
            {
                validEntries.Add(entry);
            }
            else
            {
                invalidEntries.Add(entry);
                Logger.Debug("Invalid entry: {Path}", entry.Path);
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
            var isValid = MatchesAutodeskTarget(entry.Path);
            
            if (isValid)
            {
                validEntries.Add(entry);
            }
            else
            {
                invalidCount++;
                if (invalidCount <= 3) // Log first 3 invalid entries for debugging
                {
                    Logger.Debug("FilterValidEntries: Invalid file entry {Path}", entry.Path);
                }
            }
        }
        
        Logger.Info("FilterValidEntries: Found {ValidCount} valid file system entries, {InvalidCount} invalid entries", validEntries.Count, invalidCount);
        return validEntries;
    }

    /// <summary>
    /// Collapses nested file and directory entries into the smallest safe set of removals.
    /// Prefers deleting a parent directory recursively instead of deleting its children one by one.
    /// </summary>
    /// <param name="entries">The validated file system entries.</param>
    /// <returns>The optimized removal plan.</returns>
    private static List<FileSystemEntry> BuildRemovalPlan(IReadOnlyCollection<FileSystemEntry> entries)
    {
        if (entries.Count == 0)
        {
            return [];
        }

        var orderedEntries = entries
            .Select(entry => (Entry: entry, FullPath: NormalizeFullPath(entry.Path)))
            .OrderBy(item => item.Entry.EntryType == FileSystemEntryType.Directory ? 0 : 1)
            .ThenBy(item => item.FullPath.Length)
            .ThenBy(item => item.FullPath, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var plannedEntries = new List<(FileSystemEntry Entry, string FullPath)>();

        foreach (var candidate in orderedEntries)
        {
            var isCoveredByPlannedDirectory = plannedEntries.Any(planned =>
                planned.Entry.EntryType == FileSystemEntryType.Directory &&
                IsDescendantPath(candidate.FullPath, planned.FullPath));

            if (!isCoveredByPlannedDirectory)
            {
                plannedEntries.Add(candidate);
            }
        }

        return plannedEntries.Select(item => item.Entry).ToList();
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

        if (AutodeskInstallMediaPaths.Any(basePath =>
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
    /// Determines whether a path targets Autodesk products, install media, or licensing artifacts.
    /// </summary>
    /// <param name="path">The path to evaluate.</param>
    /// <returns>True if the path should be considered for cleanup.</returns>
    private static bool MatchesAutodeskTarget(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        var fileName = Path.GetFileName(path);
        return IsAutodeskPath(path) ||
               AutodeskPattern.IsMatch(path) ||
               (!string.IsNullOrEmpty(fileName) && AutodeskPattern.IsMatch(fileName)) ||
               (!string.IsNullOrEmpty(fileName) && FlexNetPattern.IsMatch(fileName)) ||
               IsAutodeskInstallMediaPath(path);
    }

    /// <summary>
    /// Detects Autodesk installer caches and extracted 2027 media.
    /// </summary>
    /// <param name="path">The path to evaluate.</param>
    /// <returns>True if the path looks like Autodesk installer media.</returns>
    private static bool IsAutodeskInstallMediaPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        var normalizedPath = path.Replace('/', '\\');
        if (AutodeskInstallMediaPaths.Any(basePath =>
            normalizedPath.StartsWith(basePath, StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        if (!Autodesk2027InstallPattern.IsMatch(normalizedPath))
        {
            return false;
        }

        return normalizedPath.Contains(@"\Autodesk\", StringComparison.OrdinalIgnoreCase) ||
               normalizedPath.Contains(@"\ODIS\", StringComparison.OrdinalIgnoreCase) ||
               normalizedPath.Contains(@"\AdODIS\", StringComparison.OrdinalIgnoreCase) ||
               normalizedPath.Contains("setup.exe", StringComparison.OrdinalIgnoreCase) ||
               normalizedPath.Contains("installer.exe", StringComparison.OrdinalIgnoreCase) ||
               normalizedPath.Contains("bundlemanifest.xml", StringComparison.OrdinalIgnoreCase) ||
               normalizedPath.Contains("collection.xml", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Normalizes a path to a comparable full-path form.
    /// </summary>
    /// <param name="path">The path to normalize.</param>
    /// <returns>The normalized full path.</returns>
    private static string NormalizeFullPath(string path)
    {
        var fullPath = Path.GetFullPath(path);
        return fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }

    /// <summary>
    /// Determines whether a candidate path is equal to or nested under a parent directory path.
    /// </summary>
    /// <param name="candidatePath">The candidate path.</param>
    /// <param name="directoryPath">The parent directory path.</param>
    /// <returns>True when the candidate is covered by the parent directory deletion.</returns>
    private static bool IsDescendantPath(string candidatePath, string directoryPath)
    {
        if (candidatePath.Equals(directoryPath, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return candidatePath.StartsWith(
            directoryPath + Path.DirectorySeparatorChar,
            StringComparison.OrdinalIgnoreCase);
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
            if (MatchesAutodeskTarget(directoryPath))
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
                    if (MatchesAutodeskTarget(file.FullName))
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
                    if (MatchesAutodeskTarget(subDirectory.FullName))
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
                    if (MatchesAutodeskTarget(item.FullName))
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
    /// Stops Autodesk services that might lock files during deletion.
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

            Logger.Info("Stopping Autodesk services that might lock files...");
        
        foreach (var serviceName in autodeskServices)
        {
            try
            {
                Logger.Debug("Attempting to stop service: {ServiceName}", serviceName);
                
                // First try to stop the service normally
                var stopResult = await StopServiceAsync(serviceName);
                if (stopResult)
                {
                    Logger.Info("Successfully stopped service: {ServiceName}", serviceName);
                    continue;
                }
                
                // If normal stop failed, try to force stop
                Logger.Debug("Normal stop failed, attempting to force stop service: {ServiceName}", serviceName);
                var forceStopResult = await ForceStopServiceAsync(serviceName);
                if (forceStopResult)
                {
                    Logger.Info("Successfully force-stopped service: {ServiceName}", serviceName);
                }
                else
                {
                    Logger.Debug("Service {ServiceName} was not running or could not be stopped", serviceName);
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
    /// Kills all Autodesk-related processes that might be locking files.
    /// </summary>
    private async Task KillAllAutodeskProcessesAsync()
    {
        Logger.Info("Killing all Autodesk-related processes that might lock files...");
        
        var autodeskProcessNames = new[] 
        {
            "adskflex", "lmgrd", "lmadmin", "AutodeskDesktopApp", "AdskLicensingAgent",
            "AdskLicensingService", "AdskAccessCore", "AdskAccessService", "AdAppMgrSvc",
            "maya", "3dsmax", "autocad", "revit", "inventor", "fusion360", "navisworks",
            "mudbox", "motionbuilder", "alias", "adsk", "autodesk"
        };
        
        var killedProcesses = 0;
        
        foreach (var processName in autodeskProcessNames)
        {
            try
            {
                var processes = Process.GetProcessesByName(processName);
                foreach (var process in processes)
                {
                    try
                    {
                        if (ProtectedProcesses.Contains(process.ProcessName))
                        {
                            Logger.Debug("Skipping protected system process: {ProcessName} (PID: {ProcessId})", 
                                process.ProcessName, process.Id);
                            continue;
                        }
                        
                        Logger.Info("Killing Autodesk process: {ProcessName} (PID: {ProcessId})", 
                            process.ProcessName, process.Id);
                        
                        if (!process.HasExited)
                        {
                            process.Kill();
                            await process.WaitForExitAsync();
                            killedProcesses++;
                            
                            Logger.Info("Successfully killed process: {ProcessName} (PID: {ProcessId})", 
                                process.ProcessName, process.Id);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Debug(ex, "Could not kill process {ProcessName} (PID: {ProcessId}): {Message}", 
                            process.ProcessName, process.Id, ex.Message);
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Debug(ex, "Error searching for processes with name {ProcessName}: {Message}", processName, ex.Message);
            }
        }
        
        if (killedProcesses > 0)
        {
            Logger.Info("Killed {KilledCount} Autodesk processes", killedProcesses);
            // Wait for processes to fully terminate and release file handles
            await Task.Delay(3000);
        }
        else
        {
            Logger.Debug("No Autodesk processes found to kill");
        }
        
        Logger.Info("Finished killing Autodesk processes");
    }
    
    /// <summary>
    /// Attempts to stop a Windows service normally using ServiceController.
    /// </summary>
    /// <param name="serviceName">The name of the service to stop.</param>
    /// <returns>True if the service was stopped successfully, false otherwise.</returns>
    private async Task<bool> StopServiceAsync(string serviceName)
    {
        try
        {
            using var service = new ServiceController(serviceName);
            
            // Check if service exists and is running
            if (service.Status == ServiceControllerStatus.Stopped || 
                service.Status == ServiceControllerStatus.StopPending)
            {
                Logger.Debug("Service {ServiceName} is already stopped or stopping", serviceName);
                return true;
            }
            
            if (service.Status != ServiceControllerStatus.Running)
            {
                Logger.Debug("Service {ServiceName} is in state {Status}, cannot stop", serviceName, service.Status);
                return false;
            }
            
            Logger.Debug("Stopping service {ServiceName} (current status: {Status})", serviceName, service.Status);
            service.Stop();
            
            // Wait for the service to stop with timeout
            var timeout = TimeSpan.FromSeconds(30);
            await Task.Run(() => service.WaitForStatus(ServiceControllerStatus.Stopped, timeout));
            
            Logger.Info("Successfully stopped service: {ServiceName}", serviceName);
            return true;
        }
        catch (InvalidOperationException ex)
        {
            Logger.Debug(ex, "Service {ServiceName} does not exist or cannot be accessed: {Message}", serviceName, ex.Message);
            return false;
        }
        catch (System.ServiceProcess.TimeoutException ex)
        {
            Logger.Debug(ex, "Timeout stopping service {ServiceName}: {Message}", serviceName, ex.Message);
            return false;
        }
        catch (Exception ex)
        {
            Logger.Debug(ex, "Error stopping service {ServiceName}: {Message}", serviceName, ex.Message);
            return false;
        }
    }
    
    /// <summary>
    /// Force stops a Windows service by killing its process.
    /// </summary>
    /// <param name="serviceName">The name of the service to force stop.</param>
    /// <returns>True if the service was force-stopped successfully, false otherwise.</returns>
    private async Task<bool> ForceStopServiceAsync(string serviceName)
    {
        try
        {
            // First, try to get the service process ID
            var serviceProcessId = await GetServiceProcessIdAsync(serviceName);
            if (serviceProcessId > 0)
            {
                try
                {
                    var serviceProcess = Process.GetProcessById(serviceProcessId);
                    if (!serviceProcess.HasExited)
                    {
                        Logger.Debug("Force killing service process {ServiceName} (PID: {ProcessId})", serviceName, serviceProcessId);
                        serviceProcess.Kill();
                        await serviceProcess.WaitForExitAsync();
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Debug(ex, "Error force killing service process {ServiceName}: {Message}", serviceName, ex.Message);
                }
            }
            
            // If we can't get the process ID, try taskkill
            var startInfo = new ProcessStartInfo
            {
                FileName = "taskkill.exe",
                Arguments = $"/F /IM {serviceName}.exe",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using var process = Process.Start(startInfo);
            if (process != null)
            {
                await process.WaitForExitAsync();
                return process.ExitCode == 0;
            }
            
            return false;
        }
        catch (Exception ex)
        {
            Logger.Debug(ex, "Error force stopping service {ServiceName}: {Message}", serviceName, ex.Message);
            return false;
        }
    }
    
    /// <summary>
    /// Checks if a Windows service is stopped.
    /// </summary>
    /// <param name="serviceName">The name of the service to check.</param>
    /// <returns>True if the service is stopped, false otherwise.</returns>
    private async Task<bool> IsServiceStoppedAsync(string serviceName)
    {
        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "sc.exe",
                Arguments = $"query \"{serviceName}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using var process = Process.Start(startInfo);
            if (process != null)
            {
                await process.WaitForExitAsync();
                var output = await process.StandardOutput.ReadToEndAsync();
                
                // Check if the service state indicates it's stopped
                return output.Contains("STOPPED", StringComparison.OrdinalIgnoreCase) ||
                       output.Contains("does not exist", StringComparison.OrdinalIgnoreCase);
            }
            return false;
        }
        catch (Exception ex)
        {
            Logger.Debug(ex, "Error checking service status {ServiceName}: {Message}", serviceName, ex.Message);
            return false;
        }
    }
    
    /// <summary>
    /// Gets the process ID of a Windows service.
    /// </summary>
    /// <param name="serviceName">The name of the service.</param>
    /// <returns>The process ID of the service, or 0 if not found.</returns>
    private async Task<int> GetServiceProcessIdAsync(string serviceName)
    {
        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "tasklist.exe",
                Arguments = $"/SVC /FI \"SERVICES eq {serviceName}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using var process = Process.Start(startInfo);
            if (process != null)
            {
                await process.WaitForExitAsync();
                var output = await process.StandardOutput.ReadToEndAsync();
                
                // Parse the output to extract the PID
                var lines = output.Split('\n', StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    if (line.Contains(serviceName, StringComparison.OrdinalIgnoreCase))
                    {
                        var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length >= 2 && int.TryParse(parts[1], out var pid))
                        {
                            return pid;
                        }
                    }
                }
            }
            return 0;
        }
        catch (Exception ex)
        {
            Logger.Debug(ex, "Error getting service process ID {ServiceName}: {Message}", serviceName, ex.Message);
            return 0;
        }
    }

    /// <summary>
    /// Kills processes that are using the specified file, excluding protected system processes.
    /// </summary>
    /// <param name="filePath">The file path to check for process usage.</param>
    /// <returns>True if processes were killed or no processes were using the file, false otherwise.</returns>
    private async Task<bool> KillProcessesUsingFileAsync(string filePath)
    {
        try
        {
            Logger.Debug("Checking for processes using file: {FilePath}", filePath);
            
            var processesUsingFile = await GetProcessesUsingFileAsync(filePath);
            if (!processesUsingFile.Any())
            {
                Logger.Debug("No processes found using file: {FilePath}", filePath);
                return true;
            }

            Logger.Info("Found {ProcessCount} processes using file: {FilePath}", processesUsingFile.Count, filePath);
            
            var killedProcesses = 0;
            foreach (var process in processesUsingFile)
            {
                try
                {
                    if (process.Id == Environment.ProcessId)
                    {
                        Logger.Warn("Skipping current cleaner process while resolving file lock: {ProcessName} (PID: {ProcessId})",
                            process.ProcessName, process.Id);
                        continue;
                    }

                    if (ProtectedProcesses.Contains(process.ProcessName))
                    {
                        Logger.Debug("Skipping protected system process: {ProcessName} (PID: {ProcessId})", 
                            process.ProcessName, process.Id);
                        continue;
                    }

                    Logger.Info("Killing process {ProcessName} (PID: {ProcessId}) that is using file: {FilePath}", 
                        process.ProcessName, process.Id, filePath);
                    
                    if (!process.HasExited)
                    {
                        process.Kill();
                        await process.WaitForExitAsync();
                    }
                    killedProcesses++;
                    
                    Logger.Info("Successfully killed process {ProcessName} (PID: {ProcessId})", 
                        process.ProcessName, process.Id);
                }
                catch (Exception ex)
                {
                    Logger.Debug(ex, "Could not kill process {ProcessName} (PID: {ProcessId}): {Message}", 
                        process.ProcessName, process.Id, ex.Message);
                }
                finally
                {
                    process.Dispose();
                }
            }
            
            if (killedProcesses > 0)
            {
                // Wait a moment for file handles to be released
                await Task.Delay(1000);
            }
            
            Logger.Debug("Killed {KilledCount} out of {TotalCount} processes using file: {FilePath}", 
                killedProcesses, processesUsingFile.Count, filePath);
            
            return true;
        }
        catch (Exception ex)
        {
            Logger.Debug(ex, "Error while killing processes using file {FilePath}: {Message}", filePath, ex.Message);
            return false;
        }
    }

    /// <summary>
    /// Gets a list of processes that are currently using the specified file.
    /// </summary>
    /// <param name="filePath">The file path to check.</param>
    /// <returns>A list of processes using the file.</returns>
    private async Task<List<Process>> GetProcessesUsingFileAsync(string filePath)
    {
        var processesUsingFile = new List<Process>();
        
        try
        {
            // Use handle.exe if available (SysInternals tool)
            var handleOutput = await RunHandleToolAsync(filePath);
            if (!string.IsNullOrEmpty(handleOutput))
            {
                return ParseHandleOutput(handleOutput);
            }

            Logger.Debug("Skipping heuristic WMI lock detection for {FilePath} because handle.exe is unavailable", filePath);
            return processesUsingFile;
        }
        catch (Exception ex)
        {
            Logger.Debug(ex, "Error getting processes using file {FilePath}: {Message}", filePath, ex.Message);
            return processesUsingFile;
        }
    }

    /// <summary>
    /// Runs the handle.exe tool to find processes using a file.
    /// </summary>
    /// <param name="filePath">The file path to check.</param>
    /// <returns>The output from handle.exe or null if not available.</returns>
    private async Task<string?> RunHandleToolAsync(string filePath)
    {
        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "handle.exe",
                Arguments = $"-nobanner \"{filePath}\"",
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
                    return await process.StandardOutput.ReadToEndAsync();
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Debug(ex, "handle.exe not available or failed: {Message}", ex.Message);
        }
        
        return null;
    }

    /// <summary>
    /// Parses the output from handle.exe to extract process information.
    /// </summary>
    /// <param name="handleOutput">The output from handle.exe.</param>
    /// <returns>A list of processes using the file.</returns>
    private static List<Process> ParseHandleOutput(string handleOutput)
    {
        var processes = new List<Process>();
        
        try
        {
            var lines = handleOutput.Split('\n', StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
            {
                // Parse lines like: "processname.exe pid: 1234 type: File"
                var parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 3 && parts[1].StartsWith("pid:"))
                {
                    if (int.TryParse(parts[1].Substring(4), out var pid))
                    {
                        try
                        {
                            var process = Process.GetProcessById(pid);
                            processes.Add(process);
                        }
                        catch
                        {
                            // Process may have exited
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Note: Using static reference since this is a static method
            LogManager.GetCurrentClassLogger().Debug(ex, "Error parsing handle.exe output: {Message}", ex.Message);
        }
        
        return processes;
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
    /// <summary>
    /// <param name="entry">The file system entry to remove.</param>
    /// <returns>True if removal was successful, false otherwise.</returns>
    private async Task<bool> RemoveFileSystemEntryAsync(FileSystemEntry entry)
    {
        try
        {
            if (entry.EntryType == FileSystemEntryType.File)
            {
                return await RemoveFileAsync(entry.Path, entry.IsReadOnly);
            }
            else if (entry.EntryType == FileSystemEntryType.Directory)
            {
                return await RemoveDirectoryAsync(entry.Path);
            }

            return false;
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Exception removing file system entry {Path}", entry.Path);
            _errors.Add($"Error removing file system entry {entry.Path}: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Attempts to remove a file, recovering ownership and ACLs if access is blocked.
    /// </summary>
    private async Task<bool> RemoveFileAsync(string filePath, bool isReadOnly)
    {
        try
        {
            if (!File.Exists(filePath))
            {
                return true;
            }

            ClearReadOnlyAttribute(filePath, isReadOnly);
            File.Delete(filePath);
            return true;
        }
        catch (Exception ex) when (IsAccessOrLockException(ex))
        {
            Logger.Warn(ex, "Recovering locked or inaccessible file before deletion: {Path}", filePath);
            return await RecoverAndDeleteFileAsync(filePath, isReadOnly);
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Error removing file {Path}: {Message}", filePath, ex.Message);
            return false;
        }
    }

    /// <summary>
    /// Performs lock recovery, ownership takeover, and ACL repair before deleting a file.
    /// </summary>
    /// <param name="filePath">The file to remove.</param>
    /// <param name="isReadOnly">Whether the file was originally marked read-only.</param>
    /// <returns>True if the file was removed successfully.</returns>
    private async Task<bool> RecoverAndDeleteFileAsync(string filePath, bool isReadOnly)
    {
        await KillProcessesUsingFileAsync(filePath);
        await TakeOwnershipAndGrantFullControlAsync(filePath, isDirectory: false);

        try
        {
            if (!File.Exists(filePath))
            {
                return true;
            }

            ClearReadOnlyAttribute(filePath, isReadOnly);
            File.Delete(filePath);
            return true;
        }
        catch (Exception retryEx)
        {
            Logger.Debug(retryEx, "File delete retry failed after ownership recovery for {Path}", filePath);
            Logger.Debug("Attempting WMI file deletion for {Path}", filePath);
            return await DeleteFileWmiAsync(filePath);
        }
    }

    /// <summary>
    /// Deletes a file using WMI.
    /// <summary>
    private Task<bool> DeleteFileWmiAsync(string filePath)
    {
        return Task.Run(() => DeleteFileWmi(filePath));
    }

    /// <summary>
    /// Deletes a file using WMI synchronously.
    /// </summary>
    private bool DeleteFileWmi(string filePath)
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "root\\cimv2",
                $"SELECT * FROM CIM_DataFile WHERE Name='{filePath.Replace("\\", "\\\\")}'")
            {
                Options = { Timeout = TimeSpan.FromSeconds(30) }
            };

            using var collection = searcher.Get();
            foreach (ManagementObject file in collection)
            {
                using (file)
                {
                    var result = file.InvokeMethod("Delete", null);
                    if (result != null && Convert.ToInt32(result) == 0)
                    {
                        Logger.Info("Successfully deleted file via WMI: {Path}", filePath);
                        return true;
                    }
                    else
                    {
                        Logger.Warn("WMI file deletion failed for {Path} with return value: {ReturnValue}",
                            filePath, result);
                    }
                }
            }
            return false;
        }
        catch (Exception ex)
        {
            Logger.Debug(ex, "WMI deletion error for file {Path}: {Message}", filePath, ex.Message);
            return false;
        }
    }

    /// <summary>
    /// Attempts to remove a directory, recovering ownership and ACLs when required.
    /// </summary>
    private async Task<bool> RemoveDirectoryAsync(string directoryPath)
    {
        if (!Directory.Exists(directoryPath))
        {
            return true;
        }

        try
        {
            return await RunWithHeartbeatAsync(
                $"Deleting directory tree {directoryPath}",
                async () =>
                {
                    ClearDirectoryAttributes(directoryPath);
                    await ForceDeleteDirectoryContentsAsync(directoryPath);
                    return !Directory.Exists(directoryPath);
                });
        }
        catch (Exception ex) when (IsAccessOrLockException(ex))
        {
            Logger.Warn(ex, "Recovering locked or inaccessible directory before deletion: {Path}", directoryPath);
            await KillAllAutodeskProcessesAsync();
            await TakeOwnershipAndGrantFullControlAsync(directoryPath, isDirectory: true);
            ClearDirectoryAttributes(directoryPath);

            try
            {
                await RunWithHeartbeatAsync(
                    $"Force deleting directory tree {directoryPath}",
                    async () =>
                    {
                        await ForceDeleteDirectoryContentsAsync(directoryPath);
                        return true;
                    });
                return !Directory.Exists(directoryPath);
            }
            catch (Exception retryEx)
            {
                Logger.Error(retryEx, "Directory delete retry failed after ownership recovery for {Path}", directoryPath);
                return false;
            }
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Error removing directory {Path}: {Message}", directoryPath, ex.Message);
            return false;
        }
    }

    /// <summary>
    /// Clears the read-only bit from a file when present.
    /// </summary>
    /// <param name="filePath">The file path.</param>
    /// <param name="wasReadOnly">Whether the file was reported as read-only.</param>
    private static void ClearReadOnlyAttribute(string filePath, bool wasReadOnly)
    {
        if (!File.Exists(filePath))
        {
            return;
        }

        var attributes = File.GetAttributes(filePath);
        if (wasReadOnly || attributes.HasFlag(FileAttributes.ReadOnly))
        {
            File.SetAttributes(filePath, attributes & ~FileAttributes.ReadOnly);
        }
    }

    /// <summary>
    /// Clears restrictive attributes from all files and directories within a directory tree.
    /// </summary>
    /// <param name="directoryPath">The directory path.</param>
    private static void ClearDirectoryAttributes(string directoryPath)
    {
        if (!Directory.Exists(directoryPath))
        {
            return;
        }

        foreach (var filePath in Directory.EnumerateFiles(directoryPath, "*", RecursiveEnumerationOptions))
        {
            try
            {
                ClearReadOnlyAttribute(filePath, wasReadOnly: false);
            }
            catch
            {
                // Best effort before ownership recovery.
            }
        }

        foreach (var subDirectoryPath in Directory.EnumerateDirectories(directoryPath, "*", RecursiveEnumerationOptions))
        {
            try
            {
                var attributes = File.GetAttributes(subDirectoryPath);
                if (attributes.HasFlag(FileAttributes.ReadOnly))
                {
                    File.SetAttributes(subDirectoryPath, attributes & ~FileAttributes.ReadOnly);
                }
            }
            catch
            {
                // Best effort before ownership recovery.
            }
        }

        try
        {
            var rootAttributes = File.GetAttributes(directoryPath);
            if (rootAttributes.HasFlag(FileAttributes.ReadOnly))
            {
                File.SetAttributes(directoryPath, rootAttributes & ~FileAttributes.ReadOnly);
            }
        }
        catch
        {
            // Best effort before ownership recovery.
        }
    }

    /// <summary>
    /// Performs a depth-first delete of a directory after ownership/ACL recovery.
    /// </summary>
    /// <param name="directoryPath">The directory to remove.</param>
    private async Task ForceDeleteDirectoryContentsAsync(string directoryPath)
    {
        foreach (var filePath in Directory.EnumerateFiles(directoryPath, "*", RecursiveEnumerationOptions))
        {
            var isReadOnly = false;
            try
            {
                isReadOnly = File.GetAttributes(filePath).HasFlag(FileAttributes.ReadOnly);
            }
            catch
            {
                // Best effort only.
            }

            await RemoveFileAsync(filePath, isReadOnly);
        }

        var subDirectories = Directory
            .EnumerateDirectories(directoryPath, "*", RecursiveEnumerationOptions)
            .OrderByDescending(static path => path.Length)
            .ToList();

        foreach (var subDirectoryPath in subDirectories)
        {
            try
            {
                var attributes = File.GetAttributes(subDirectoryPath);
                if (attributes.HasFlag(FileAttributes.ReadOnly))
                {
                    File.SetAttributes(subDirectoryPath, attributes & ~FileAttributes.ReadOnly);
                }
            }
            catch
            {
                // Best effort only.
            }

            if (Directory.Exists(subDirectoryPath))
            {
                Directory.Delete(subDirectoryPath, false);
            }
        }

        if (Directory.Exists(directoryPath))
        {
            Directory.Delete(directoryPath, false);
        }
    }

    /// <summary>
    /// Attempts to take ownership and grant the local Administrators group full control.
    /// </summary>
    /// <param name="path">The file or directory path.</param>
    /// <param name="isDirectory">Whether the path is a directory.</param>
    private async Task TakeOwnershipAndGrantFullControlAsync(string path, bool isDirectory)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        var takeOwnArguments = isDirectory
            ? $"/F \"{path}\" /A /R /D Y"
            : $"/F \"{path}\" /A";
        var icaclsArguments = isDirectory
            ? $"\"{path}\" /grant *S-1-5-32-544:(OI)(CI)F /T /C /Q"
            : $"\"{path}\" /grant *S-1-5-32-544:F /C /Q";

        await RunUtilityAsync("takeown.exe", takeOwnArguments, path);
        await RunUtilityAsync("icacls.exe", icaclsArguments, path);
    }

    /// <summary>
    /// Runs a Windows utility used during delete recovery.
    /// </summary>
    /// <param name="fileName">The utility name.</param>
    /// <param name="arguments">Command-line arguments.</param>
    /// <param name="targetPath">The path being processed.</param>
    private async Task RunUtilityAsync(string fileName, string arguments, string targetPath)
    {
        try
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
                Logger.Warn("Failed to start utility {Utility} for {Path}", fileName, targetPath);
                return;
            }

            await process.WaitForExitAsync();

            if (process.ExitCode != 0)
            {
                var errorOutput = await process.StandardError.ReadToEndAsync();
                Logger.Debug("Utility {Utility} exited with code {ExitCode} for {Path}: {ErrorOutput}",
                    fileName, process.ExitCode, targetPath, errorOutput.Trim());
            }
        }
        catch (Exception ex)
        {
            Logger.Debug(ex, "Utility {Utility} failed for {Path}: {Message}", fileName, targetPath, ex.Message);
        }
    }

    /// <summary>
    /// Determines whether an exception indicates a lock or access-control problem.
    /// </summary>
    /// <param name="ex">The exception to evaluate.</param>
    /// <returns>True if the exception should trigger ownership recovery.</returns>
    private static bool IsAccessOrLockException(Exception ex)
    {
        if (ex is UnauthorizedAccessException)
        {
            return true;
        }

        if (ex is not IOException ioException)
        {
            return false;
        }

        return ioException.HResult == unchecked((int)0x80070005) ||
               ioException.HResult == unchecked((int)0x80070020) ||
               ioException.HResult == unchecked((int)0x80070021) ||
               ioException.Message.Contains("access is denied", StringComparison.OrdinalIgnoreCase) ||
               ioException.Message.Contains("being used by another process", StringComparison.OrdinalIgnoreCase) ||
               ioException.Message.Contains("cannot access the file", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Runs a long operation and periodically reports that it is still active.
    /// </summary>
    /// <typeparam name="T">The operation result type.</typeparam>
    /// <param name="description">The operation description.</param>
    /// <param name="operation">The operation to run.</param>
    /// <returns>The operation result.</returns>
    private async Task<T> RunWithHeartbeatAsync<T>(string description, Func<Task<T>> operation)
    {
        var operationTask = operation();
        var stopwatch = Stopwatch.StartNew();
        var heartbeatInterval = TimeSpan.FromSeconds(3);

        while (!operationTask.IsCompleted)
        {
            var completedTask = await Task.WhenAny(operationTask, Task.Delay(heartbeatInterval));
            if (completedTask != operationTask)
            {
                Logger.Info("{Description} is still in progress ({ElapsedSeconds:F0}s elapsed)", description, stopwatch.Elapsed.TotalSeconds);

                if (!_config.DryRun)
                {
                    AnsiConsole.MarkupLine(
                        $"[dim]{Markup.Escape(description)} still in progress ({stopwatch.Elapsed.TotalSeconds:F0}s elapsed)...[/]");
                }
            }
        }

        return await operationTask;
    }

    /// <summary>
    /// Formats an estimated remaining duration from elapsed work and completed item count.
    /// </summary>
    /// <param name="elapsed">Elapsed processing time.</param>
    /// <param name="processedCount">Completed item count.</param>
    /// <param name="totalCount">Total item count.</param>
    /// <returns>Human-readable ETA text.</returns>
    private static string FormatEta(TimeSpan elapsed, int processedCount, int totalCount)
    {
        if (processedCount <= 0 || totalCount <= processedCount)
        {
            return processedCount >= totalCount ? "done" : "calculating...";
        }

        var averageSecondsPerItem = elapsed.TotalSeconds / processedCount;
        var remainingSeconds = Math.Max(0, (totalCount - processedCount) * averageSecondsPerItem);
        var remaining = TimeSpan.FromSeconds(remainingSeconds);

        if (remaining.TotalHours >= 1)
        {
            return $"{(int)remaining.TotalHours}h {remaining.Minutes}m";
        }

        if (remaining.TotalMinutes >= 1)
        {
            return $"{remaining.Minutes}m {remaining.Seconds}s";
        }

        return $"{Math.Max(1, remaining.Seconds)}s";
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
