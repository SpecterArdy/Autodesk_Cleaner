using System.Diagnostics;
using System.IO.Compression;
using System.IO.Abstractions;
using System.Management;
using System.Runtime.InteropServices;
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

        // Stop Autodesk services before deletion
        await StopAutodeskServicesAsync();
        
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

                Logger.Debug("Removing file system entry: {Path}", entry.Path);
                if (await RemoveFileSystemEntryAsync(entry))
                {
                    Logger.Debug("Successfully removed file system entry: {Path}", entry.Path);
                    successfulRemovals++;
                }
                else
                {
                    Logger.Warn("Failed to remove file system entry: {Path}", entry.Path);
                    errors.Add($"Failed to remove: {entry.Path}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Error removing file system entry {Path}: {Message}", entry.Path, ex.Message);
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
                    if (ProtectedProcesses.Contains(process.ProcessName))
                    {
                        Logger.Debug("Skipping protected system process: {ProcessName} (PID: {ProcessId})", 
                            process.ProcessName, process.Id);
                        continue;
                    }

                    Logger.Info("Killing process {ProcessName} (PID: {ProcessId}) that is using file: {FilePath}", 
                        process.ProcessName, process.Id, filePath);
                    
                    process.Kill();
                    await process.WaitForExitAsync();
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
            
            // Fallback: Use WMI to find processes with open file handles
            return await GetProcessesUsingFileWmiAsync(filePath);
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
    /// Uses WMI to find processes that might be using a file (fallback method).
    /// </summary>
    /// <param name="filePath">The file path to check.</param>
    /// <returns>A list of processes that might be using the file.</returns>
    private async Task<List<Process>> GetProcessesUsingFileWmiAsync(string filePath)
    {
        var processes = new List<Process>();
        
        try
        {
            await Task.Run(() =>
            {
                using var searcher = new ManagementObjectSearcher("SELECT ProcessId, Name FROM Win32_Process");
                using var results = searcher.Get();
                
                foreach (ManagementObject result in results)
                {
                    try
                    {
                        var processId = Convert.ToInt32(result["ProcessId"]);
                        var processName = result["Name"]?.ToString();
                        
                        if (processName != null && !ProtectedProcesses.Contains(processName))
                        {
                            // Check if this process might be related to Autodesk or the file directory
                            var fileDirectory = Path.GetDirectoryName(filePath);
                            if (processName.Contains("autodesk", StringComparison.OrdinalIgnoreCase) ||
                                (fileDirectory != null && processName.Contains(Path.GetFileName(fileDirectory), StringComparison.OrdinalIgnoreCase)))
                            {
                                try
                                {
                                    var process = Process.GetProcessById(processId);
                                    processes.Add(process);
                                }
                                catch
                                {
                                    // Process may have exited
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Skip invalid processes
                    }
                }
            });
        }
        catch (Exception ex)
        {
            Logger.Debug(ex, "Error using WMI to find processes: {Message}", ex.Message);
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
    /// </summary>
    /// <param name="entry">The file system entry to remove.</param>
    /// <returns>True if removal was successful, false otherwise.</returns>
    private async Task<bool> RemoveFileSystemEntryAsync(FileSystemEntry entry)
    {
        try
        {
            if (entry.EntryType == FileSystemEntryType.File)
            {
                if (File.Exists(entry.Path))
                {
                    // First attempt: try to delete normally
                    try
                    {
                        // Remove read-only attribute if present
                        if (entry.IsReadOnly)
                        {
                            File.SetAttributes(entry.Path, File.GetAttributes(entry.Path) & ~FileAttributes.ReadOnly);
                        }

                        File.Delete(entry.Path);
                        return true;
                    }
                    catch (IOException ex) when (ex.Message.Contains("being used by another process") || 
                                                ex.Message.Contains("access is denied"))
                    {
                        Logger.Debug("File {Path} is in use, attempting to kill processes using it", entry.Path);
                        
                        // Kill processes using the file
                        await KillProcessesUsingFileAsync(entry.Path);
                        
                        // Try again after killing processes
                        try
                        {
                            if (entry.IsReadOnly)
                            {
                                File.SetAttributes(entry.Path, File.GetAttributes(entry.Path) & ~FileAttributes.ReadOnly);
                            }
                            File.Delete(entry.Path);
                            return true;
                        }
                        catch (Exception retryEx)
                        {
                            Logger.Debug(retryEx, "Failed to delete file {Path} even after killing processes: {Message}", 
                                entry.Path, retryEx.Message);
                            throw;
                        }
                    }
                }
            }
            else if (entry.EntryType == FileSystemEntryType.Directory)
            {
                if (Directory.Exists(entry.Path))
                {
                    // First attempt: try to delete normally
                    try
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
                        return true;
                    }
                    catch (IOException ex) when (ex.Message.Contains("being used by another process") || 
                                                ex.Message.Contains("access is denied"))
                    {
                        Logger.Debug("Directory {Path} contains files in use, attempting to kill processes", entry.Path);
                        
                        // Kill processes using files in the directory
                        var directoryInfo = new DirectoryInfo(entry.Path);
                        var filesInUse = directoryInfo.GetFiles("*", SearchOption.AllDirectories);
                        
                        foreach (var file in filesInUse)
                        {
                            await KillProcessesUsingFileAsync(file.FullName);
                        }
                        
                        // Try again after killing processes
                        try
                        {
                            // Remove read-only attributes again
                            foreach (var file in directoryInfo.GetFiles("*", SearchOption.AllDirectories))
                            {
                                if (file.IsReadOnly)
                                {
                                    file.Attributes &= ~FileAttributes.ReadOnly;
                                }
                            }
                            
                            Directory.Delete(entry.Path, true);
                            return true;
                        }
                        catch (Exception retryEx)
                        {
                            Logger.Debug(retryEx, "Failed to delete directory {Path} even after killing processes: {Message}", 
                                entry.Path, retryEx.Message);
                            throw;
                        }
                    }
                }
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
