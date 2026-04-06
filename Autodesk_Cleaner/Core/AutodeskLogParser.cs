using NLog;
using System.Diagnostics.Eventing.Reader;
using System.Text.RegularExpressions;

namespace Autodesk_Cleaner.Core;

/// <summary>
/// Parses Autodesk ODIS logs and Windows Installer event data for stale MSI failures.
/// </summary>
internal sealed class AutodeskLogParser
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
    private static readonly Regex ProductCodeRegex = new(@"\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}", RegexOptions.Compiled);
    private static readonly Regex ProductLineRegex = new(@"\b(Maya|3ds\s*Max|AutoCAD|Revit|Inventor|Navisworks|MotionBuilder|Mudbox|Alias)\s+20\d{2}\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex MainEngine1612Regex = new(@"MainEngineThread\s+is\s+returning\s+1612", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ResolveSourceRegex = new(@"Failed\s+to\s+resolve\s+source", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex SourceUnavailableRegex = new(@"The\s+installation\s+source\s+for\s+this\s+product\s+is\s+not\s+available", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex CorruptConfigurationRegex = new(@"configuration\s+data\s+for\s+this\s+product\s+is\s+corrupt", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex ReturnCodeRegex = new(@"(?<package>[^\\/:*?""<>|\r\n]+?\.msi)\s+return\s+code\s+(?<code>\d+)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex FailedPackageRegex = new(@"Failed\s+to\s+install\s+Package.*?(?<name>Autodesk Genuine Service|[^\\/:*?""<>|\r\n]+?)(?:\s{2,}|\s*$)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex MediaPathRegex = new(@"(?<path>[A-Za-z]:\\[^""\r\n]+?\.msi)", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex PackedKeyRegex = new(@"Installer\\Products\\(?<packed>[A-Za-z0-9]{32})\\SourceList", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    private static readonly Regex AutodeskPackageRegex = new(@"(autodesk genuine service|autodesk|adsk|cer|genuine service)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Finds recent Autodesk installer failures from logs and Application event log data.
    /// </summary>
    /// <returns>Parsed failure records.</returns>
    public IReadOnlyCollection<AutodeskInstallFailureRecord> FindInstallerFailures()
    {
        var results = new List<AutodeskInstallFailureRecord>();

        foreach (var logPath in EnumerateCandidateLogFiles())
        {
            try
            {
                var parsedRecord = TryParseLog(logPath);
                if (parsedRecord is not null)
                {
                    results.Add(parsedRecord.Value);
                }
            }
            catch (Exception ex)
            {
                Logger.Debug(ex, "Failed to parse Autodesk installer log {LogPath}", logPath);
            }
        }

        results.AddRange(ReadInstallerEvents());

        return results
            .OrderByDescending(result => result.LastWriteTimeUtc)
            .ToList();
    }

    private static AutodeskInstallFailureRecord? TryParseLog(string logPath)
    {
        var errorCode = 0;
        var sourceFailureDetected = false;
        var corruptConfigurationDetected = false;
        var missingPackageNameEventDetected = false;
        string? productCode = null;
        string? packageName = TryGetPackageNameFromFileName(logPath);
        string? packageDisplayName = null;
        string? mediaPath = null;
        string? productLine = null;
        var reasons = new List<string>();

        foreach (var line in File.ReadLines(logPath))
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            if (line.Contains("1612", StringComparison.OrdinalIgnoreCase))
            {
                errorCode = 1612;
            }

            var returnCodeMatch = ReturnCodeRegex.Match(line);
            if (returnCodeMatch.Success && int.TryParse(returnCodeMatch.Groups["code"].Value, out var returnCode))
            {
                errorCode = returnCode;
                packageName ??= returnCodeMatch.Groups["package"].Value.Trim();
                packageDisplayName ??= Path.GetFileNameWithoutExtension(packageName);

                if (returnCode == 1610)
                {
                    reasons.Add("MSI returned 1610");
                }
            }

            if (ResolveSourceRegex.IsMatch(line))
            {
                sourceFailureDetected = true;
                reasons.Add("Failed to resolve source");
            }

            if (MainEngine1612Regex.IsMatch(line))
            {
                sourceFailureDetected = true;
                errorCode = 1612;
                reasons.Add("MainEngineThread returned 1612");
            }

            if (SourceUnavailableRegex.IsMatch(line))
            {
                sourceFailureDetected = true;
                reasons.Add("MSI reported unavailable installation source");
            }

            if (CorruptConfigurationRegex.IsMatch(line))
            {
                corruptConfigurationDetected = true;
                if (errorCode == 0)
                {
                    errorCode = 1610;
                }

                reasons.Add("MSI reported corrupt configuration data");
            }

            var failedPackageMatch = FailedPackageRegex.Match(line);
            if (failedPackageMatch.Success)
            {
                packageDisplayName ??= failedPackageMatch.Groups["name"].Value.Trim();
                reasons.Add($"Failed to install package {packageDisplayName}");
            }

            var mediaPathMatch = MediaPathRegex.Match(line);
            if (mediaPathMatch.Success)
            {
                mediaPath ??= mediaPathMatch.Groups["path"].Value.Trim();
                packageName ??= Path.GetFileName(mediaPath);
                packageDisplayName ??= Path.GetFileNameWithoutExtension(mediaPath);
            }

            if (productCode is null)
            {
                var looseGuidMatch = ProductCodeRegex.Match(line);
                if (looseGuidMatch.Success)
                {
                    productCode = looseGuidMatch.Value;
                }
            }

            if (productLine is null)
            {
                var productLineMatch = ProductLineRegex.Match(line);
                if (productLineMatch.Success)
                {
                    productLine = productLineMatch.Value;
                }
            }

            if (packageDisplayName is null &&
                line.Contains("Autodesk Genuine Service", StringComparison.OrdinalIgnoreCase))
            {
                packageDisplayName = "Autodesk Genuine Service";
            }
        }

        if (!sourceFailureDetected && !corruptConfigurationDetected && errorCode is not 1610 and not 1612)
        {
            return null;
        }

        if (productLine is null)
        {
            productLine = TryGetProductLineFromPath(logPath);
        }

        var failureReason = reasons.Count > 0
            ? string.Join("; ", reasons.Distinct(StringComparer.OrdinalIgnoreCase))
            : $"MSI error {errorCode} detected";

        return new AutodeskInstallFailureRecord(
            LogPath: logPath,
            PackageName: packageName,
            PackageDisplayName: packageDisplayName ?? packageName,
            ProductCode: productCode,
            PackedProductCode: null,
            ErrorCode: errorCode == 0 ? 1612 : errorCode,
            FailureReason: failureReason,
            SourceFailureDetected: sourceFailureDetected,
            CorruptConfigurationDetected: corruptConfigurationDetected,
            MissingPackageNameEventDetected: missingPackageNameEventDetected,
            MediaPath: mediaPath,
            ProductLine: productLine,
            LastWriteTimeUtc: File.GetLastWriteTimeUtc(logPath));
    }

    private static IReadOnlyCollection<AutodeskInstallFailureRecord> ReadInstallerEvents()
    {
        var results = new List<AutodeskInstallFailureRecord>();

        try
        {
            var query = new EventLogQuery(
                "Application",
                PathType.LogName,
                "*[System[Provider[@Name='MsiInstaller'] and (EventID=1002)]]")
            {
                ReverseDirection = true
            };

            using var reader = new EventLogReader(query);
            for (var index = 0; index < 200; index++)
            {
                using var eventRecord = reader.ReadEvent();
                if (eventRecord is null)
                {
                    break;
                }

                var description = eventRecord.FormatDescription();
                if (string.IsNullOrWhiteSpace(description) ||
                    !description.Contains("PackageName", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var packedKeyMatch = PackedKeyRegex.Match(description);
                if (!packedKeyMatch.Success)
                {
                    continue;
                }

                var failureReason = "Event Viewer reported missing PackageName in SourceList";
                var packageDisplayName = description.Contains("Autodesk Genuine Service", StringComparison.OrdinalIgnoreCase)
                    ? "Autodesk Genuine Service"
                    : "Windows Installer Event 1002";

                results.Add(new AutodeskInstallFailureRecord(
                    LogPath: "Application Event Log",
                    PackageName: null,
                    PackageDisplayName: packageDisplayName,
                    ProductCode: null,
                    PackedProductCode: packedKeyMatch.Groups["packed"].Value,
                    ErrorCode: 1610,
                    FailureReason: failureReason,
                    SourceFailureDetected: false,
                    CorruptConfigurationDetected: false,
                    MissingPackageNameEventDetected: true,
                    MediaPath: null,
                    ProductLine: null,
                    LastWriteTimeUtc: eventRecord.TimeCreated?.ToUniversalTime() ?? DateTime.UtcNow));
            }
        }
        catch (Exception ex)
        {
            Logger.Debug(ex, "Failed to read MsiInstaller Application events");
        }

        return results;
    }

    private static IEnumerable<string> EnumerateCandidateLogFiles()
    {
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var path in GetPreferredLogPaths())
        {
            if (File.Exists(path) && visited.Add(path))
            {
                yield return path;
            }
        }

        foreach (var directoryPath in GetCandidateDirectories())
        {
            if (!Directory.Exists(directoryPath))
            {
                continue;
            }

            foreach (var logPath in EnumerateLogs(directoryPath))
            {
                if (visited.Add(logPath))
                {
                    yield return logPath;
                }
            }
        }
    }

    private static IEnumerable<string> GetPreferredLogPaths()
    {
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        yield return Path.Combine(localAppData, "Autodesk", "ODIS", "Install.log");
        yield return Path.Combine(localAppData, "Autodesk", "ODIS", "Setup.log");
    }

    private static IEnumerable<string> EnumerateLogs(string rootDirectory)
    {
        IEnumerable<string> filePaths;

        try
        {
            filePaths = Directory.EnumerateFiles(rootDirectory, "*.log", SearchOption.AllDirectories);
        }
        catch
        {
            yield break;
        }

        foreach (var filePath in filePaths
                     .Where(path =>
                         path.Contains("ODIS", StringComparison.OrdinalIgnoreCase) ||
                         path.Contains("_install", StringComparison.OrdinalIgnoreCase) ||
                         path.Contains("setup", StringComparison.OrdinalIgnoreCase) ||
                         AutodeskPackageRegex.IsMatch(path))
                     .OrderByDescending(path => File.GetLastWriteTimeUtc(path))
                     .Take(250))
        {
            yield return filePath;
        }
    }

    private static IEnumerable<string> GetCandidateDirectories()
    {
        yield return @"C:\Autodesk";
        yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Autodesk");
        yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Autodesk");

        var tempPath = Path.GetTempPath();
        if (!string.IsNullOrWhiteSpace(tempPath))
        {
            yield return tempPath;
        }
    }

    private static string? TryGetPackageNameFromFileName(string logPath)
    {
        var fileName = Path.GetFileNameWithoutExtension(logPath);
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return null;
        }

        return fileName.EndsWith("_install", StringComparison.OrdinalIgnoreCase)
            ? fileName[..^"_install".Length]
            : fileName;
    }

    private static string? TryGetProductLineFromPath(string path)
    {
        var match = ProductLineRegex.Match(path);
        return match.Success ? match.Value : null;
    }
}

/// <summary>
/// A parsed Autodesk install failure record.
/// </summary>
/// <param name="LogPath">The source log path.</param>
/// <param name="PackageName">The package name when identified.</param>
/// <param name="PackageDisplayName">The package display name when identified.</param>
/// <param name="ProductCode">The MSI product code when identified.</param>
/// <param name="PackedProductCode">The packed product code when identified.</param>
/// <param name="ErrorCode">The detected MSI error code.</param>
/// <param name="FailureReason">The parsed failure reason.</param>
/// <param name="SourceFailureDetected">Whether the failure was source-resolution related.</param>
/// <param name="CorruptConfigurationDetected">Whether the failure indicates corrupt MSI configuration data.</param>
/// <param name="MissingPackageNameEventDetected">Whether the Application event log reported missing PackageName.</param>
/// <param name="MediaPath">The media path when identified.</param>
/// <param name="ProductLine">The Autodesk product line, such as Maya 2027.</param>
/// <param name="LastWriteTimeUtc">The last write time of the source record.</param>
internal readonly record struct AutodeskInstallFailureRecord(
    string LogPath,
    string? PackageName,
    string? PackageDisplayName,
    string? ProductCode,
    string? PackedProductCode,
    int ErrorCode,
    string FailureReason,
    bool SourceFailureDetected,
    bool CorruptConfigurationDetected,
    bool MissingPackageNameEventDetected,
    string? MediaPath,
    string? ProductLine,
    DateTime LastWriteTimeUtc);
