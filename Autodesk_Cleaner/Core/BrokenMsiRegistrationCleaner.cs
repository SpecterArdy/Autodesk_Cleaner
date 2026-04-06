using Microsoft.Win32;
using NLog;
using Spectre.Console;
using System.Management;
using System.Text.RegularExpressions;

namespace Autodesk_Cleaner.Core;

/// <summary>
/// Finds and removes stale Autodesk-related Windows Installer registrations.
/// </summary>
internal sealed class BrokenMsiRegistrationCleaner
{
    private const string ProductsRoot = @"SOFTWARE\Classes\Installer\Products";
    private const string FeaturesRoot = @"SOFTWARE\Classes\Installer\Features";
    private const string UserDataProductsRoot = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products";
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
    private static readonly Regex AutodeskPattern = new(
        @"(autodesk|adsk|3ds\s*max|maya|revit|inventor|navisworks|motionbuilder|mudbox|genuine|odis|cer)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly ScannerConfig _config;

    public BrokenMsiRegistrationCleaner(ScannerConfig config)
    {
        _config = config;
    }

    /// <summary>
    /// Detects and removes Autodesk MSI registrations that should be purged during cleanup.
    /// </summary>
    /// <returns>The cleanup result.</returns>
    public async Task<BrokenMsiCleanupResult> CleanupAsync()
    {
        AnsiConsole.MarkupLine("[yellow]Enabling registry ownership privileges for installer cleanup...[/]");
        RegistryAclHelper.EnableRequiredPrivileges(Logger);

        AnsiConsole.MarkupLine(@"[yellow]Scanning HKLM\SOFTWARE\Classes\Installer\Products for Autodesk MSI registrations...[/]");
        var installerProductCount = GetInstallerProductCount();
        AnsiConsole.MarkupLine($"[green]Found {installerProductCount} Windows Installer product registrations to inspect.[/]");

        AnsiConsole.MarkupLine(
            @"[yellow]Checking ProductName, Publisher, SourceList\PackageName, SourceList\LastUsedSource, media-only SourceList entries, and UserData completeness...[/]");
        var candidates = FindCandidates();
        AnsiConsole.MarkupLine($"[green]Identified {candidates.Count} stale Autodesk MSI registration candidates.[/]");

        if (candidates.Count > 0)
        {
            foreach (var candidate in candidates.Take(10))
            {
                var reasonSummary = string.Join("; ", candidate.Reasons.Take(2));
                if (string.IsNullOrWhiteSpace(reasonSummary))
                {
                    reasonSummary = "Autodesk installer metadata is stale";
                }

                AnsiConsole.MarkupLine(
                    $"[dim]Candidate:[/] {Markup.Escape(GetCandidateLabel(candidate))} [dim]({Markup.Escape(candidate.PackedProductCode)})[/] [grey]- {Markup.Escape(reasonSummary)}[/]");
            }

            if (candidates.Count > 10)
            {
                AnsiConsole.MarkupLine($"[dim]... and {candidates.Count - 10} more stale installer registrations.[/]");
            }
        }

        AnsiConsole.MarkupLine("[yellow]Checking for running msiexec processes left behind by Autodesk installs...[/]");
        var runningMsiexec = DetectRunningMsiexec();
        if (runningMsiexec.Count > 0)
        {
            AnsiConsole.MarkupLine($"[yellow]Found {runningMsiexec.Count} running msiexec process(es).[/]");
        }
        else
        {
            AnsiConsole.MarkupLine("[green]No running msiexec processes detected.[/]");
        }

        var errors = new List<string>();
        var actionCount = 0;

        Logger.Info("Autodesk MSI cleanup candidates identified: {@Candidates}", candidates);

        foreach (var (candidate, index) in candidates.Select((candidate, index) => (candidate, index + 1)))
        {
            if (_config.DryRun)
            {
                AnsiConsole.MarkupLine(
                    $"[dim][DRY RUN] Would remove stale MSI registration {index}/{candidates.Count}:[/] {Markup.Escape(GetCandidateLabel(candidate))} [dim]({Markup.Escape(candidate.PackedProductCode)})[/]");
                actionCount++;
                continue;
            }

            if (_config.CreateBackup)
            {
                AnsiConsole.MarkupLine(
                    $"[yellow]Backing up installer registration {index}/{candidates.Count}:[/] {Markup.Escape(GetCandidateLabel(candidate))} [dim]({Markup.Escape(candidate.PackedProductCode)})[/]");
                var backupErrors = await BackupCandidateAsync(candidate);
                errors.AddRange(backupErrors);
            }

            AnsiConsole.MarkupLine(
                $"[yellow]Removing stale installer registration {index}/{candidates.Count}:[/] {Markup.Escape(GetCandidateLabel(candidate))} [dim]({Markup.Escape(candidate.PackedProductCode)})[/]");

            if (RemoveCandidate(candidate, out var error))
            {
                AnsiConsole.MarkupLine(
                    $"[green]Removed stale installer registration:[/] {Markup.Escape(GetCandidateLabel(candidate))} [dim]({Markup.Escape(candidate.PackedProductCode)})[/]");
                actionCount++;
            }
            else
            {
                AnsiConsole.MarkupLine(
                    $"[red]Failed to remove stale installer registration:[/] {Markup.Escape(GetCandidateLabel(candidate))} [dim]({Markup.Escape(candidate.PackedProductCode)})[/]");
                errors.Add(error ?? $"Failed to remove Autodesk MSI registration {candidate.PackedProductCode}");
            }
        }

        return new BrokenMsiCleanupResult(
            Candidates: candidates,
            RunningMsiexec: runningMsiexec,
            SuccessfulRemovals: actionCount,
            Errors: errors);
    }

    private List<AutodeskMsiRegistrationCandidate> FindCandidates()
    {
        var candidateMap = new Dictionary<string, CandidateBuilder>(StringComparer.OrdinalIgnoreCase);

        using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
        using var productsKey = baseKey.OpenSubKey(ProductsRoot, writable: false);
        if (productsKey is null)
        {
            return [];
        }

        foreach (var packedProductCode in productsKey.GetSubKeyNames())
        {
            var candidate = InspectCandidate(baseKey, packedProductCode, candidateMap);
            if (candidate is null)
            {
                continue;
            }

            candidateMap[packedProductCode] = candidate;
        }

        return candidateMap.Values
            .Where(builder => builder.IsStale && builder.IsAutodeskRelated)
            .Select(builder => builder.Build())
            .OrderBy(candidate => candidate.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(candidate => candidate.PackedProductCode, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private CandidateBuilder? InspectCandidate(
        RegistryKey baseKey,
        string packedProductCode,
        IDictionary<string, CandidateBuilder> candidateMap)
    {
        using var productKey = baseKey.OpenSubKey($@"{ProductsRoot}\{packedProductCode}", writable: false);
        using var productSourceListKey = baseKey.OpenSubKey($@"{ProductsRoot}\{packedProductCode}\SourceList", writable: false);
        using var productSourceMediaKey = baseKey.OpenSubKey($@"{ProductsRoot}\{packedProductCode}\SourceList\Media", writable: false);
        using var featureKey = baseKey.OpenSubKey($@"{FeaturesRoot}\{packedProductCode}", writable: false);
        using var userDataKey = baseKey.OpenSubKey($@"{UserDataProductsRoot}\{packedProductCode}", writable: false);
        using var userDataSourceListKey = baseKey.OpenSubKey($@"{UserDataProductsRoot}\{packedProductCode}\SourceList", writable: false);
        using var userDataSourceMediaKey = baseKey.OpenSubKey($@"{UserDataProductsRoot}\{packedProductCode}\SourceList\Media", writable: false);
        using var installPropertiesKey = baseKey.OpenSubKey($@"{UserDataProductsRoot}\{packedProductCode}\InstallProperties", writable: false);

        var rawProductName = ReadRegistryString(productKey, "ProductName") ??
                             ReadRegistryString(installPropertiesKey, "DisplayName") ??
                             ReadRegistryString(installPropertiesKey, "ProductName");
        var publisher = ReadRegistryString(productKey, "Publisher") ??
                        ReadRegistryString(installPropertiesKey, "Publisher");

        var packageName = FirstNonEmpty(
            ReadRegistryString(productSourceListKey, "PackageName"),
            ReadRegistryString(userDataSourceListKey, "PackageName"));

        var lastUsedSource = FirstNonEmpty(
            ReadRegistryString(productSourceListKey, "LastUsedSource"),
            ReadRegistryString(userDataSourceListKey, "LastUsedSource"));

        var installSource = ReadRegistryString(installPropertiesKey, "InstallSource");
        var onlyMediaEntries = HasOnlyMediaEntries(productSourceListKey, productSourceMediaKey, packageName, lastUsedSource) ||
                               HasOnlyMediaEntries(userDataSourceListKey, userDataSourceMediaKey, packageName, lastUsedSource);
        var userDataIncomplete = userDataKey is null || installPropertiesKey is null || featureKey is null;
        var isAutodeskRelated = MatchesAutodesk(rawProductName) ||
                                MatchesAutodesk(publisher) ||
                                MatchesAutodesk(packageName) ||
                                MatchesAutodesk(installSource);

        if (!isAutodeskRelated && !candidateMap.ContainsKey(packedProductCode))
        {
            return null;
        }

        var builder = GetOrCreateCandidate(candidateMap, packedProductCode);
        builder.DisplayName = string.IsNullOrWhiteSpace(rawProductName) ? packedProductCode : rawProductName;
        builder.PackageDisplayName ??= rawProductName;
        builder.Publisher = publisher;
        builder.PackageName ??= packageName;
        builder.LastUsedSource = lastUsedSource;
        builder.IsAutodeskRelated |= isAutodeskRelated;

        if (builder.IsAutodeskRelated)
        {
            builder.Reasons.Add("Autodesk Windows Installer registration found");
        }

        if (string.IsNullOrWhiteSpace(rawProductName))
        {
            builder.IsStale = true;
            builder.Reasons.Add("ProductName is blank");
        }

        if (string.IsNullOrWhiteSpace(packageName))
        {
            builder.IsStale = true;
            builder.Reasons.Add(@"SourceList\PackageName is missing or blank");
        }

        if (string.IsNullOrWhiteSpace(lastUsedSource))
        {
            builder.IsStale = true;
            builder.Reasons.Add(@"SourceList\LastUsedSource is missing");
        }
        else
        {
            var resolvedPath = ExtractSourcePath(lastUsedSource);
            if (IsTemporaryPath(resolvedPath))
            {
                builder.IsStale = true;
                builder.Reasons.Add($"LastUsedSource points to temp media: {lastUsedSource}");
            }
            else if (!PathExists(resolvedPath))
            {
                builder.IsStale = true;
                builder.Reasons.Add($"LastUsedSource path does not exist: {lastUsedSource}");
            }
        }

        if (onlyMediaEntries)
        {
            builder.IsStale = true;
            builder.Reasons.Add("SourceList only contains Media entries and no usable PackageName/LastUsedSource");
        }

        if (userDataIncomplete)
        {
            builder.IsStale = true;
            builder.Reasons.Add("UserData product registration is missing or incomplete");
        }

        return builder;
    }

    private async Task<List<string>> BackupCandidateAsync(AutodeskMsiRegistrationCandidate candidate)
    {
        var errors = new List<string>();
        var backupDirectory = Path.Combine(_config.BackupPath, "msi_registry_backup", candidate.PackedProductCode);

        foreach (var keyPath in candidate.RegistryPaths)
        {
            var regPath = $@"HKEY_LOCAL_MACHINE\{keyPath}";
            var backupFileName = keyPath
                .Replace('\\', '_')
                .Replace(':', '_')
                + ".reg";
            var backupFilePath = Path.Combine(backupDirectory, backupFileName);

            if (!await RegistryAclHelper.ExportKeyAsync(regPath, backupFilePath, Logger))
            {
                errors.Add($"Failed to back up registry key: {regPath}");
            }
        }

        return errors;
    }

    private bool RemoveCandidate(AutodeskMsiRegistrationCandidate candidate, out string? error)
    {
        error = null;

        foreach (var keyPath in candidate.RegistryPaths)
        {
            if (RegistryAclHelper.DeleteTree(RegistryHive.LocalMachine, RegistryView.Default, keyPath, Logger))
            {
                continue;
            }

            error = $"Failed to remove Autodesk MSI registration key HKLM\\{keyPath}";
            return false;
        }

        return true;
    }

    private static int GetInstallerProductCount()
    {
        try
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default);
            using var productsKey = baseKey.OpenSubKey(ProductsRoot, writable: false);
            return productsKey?.SubKeyCount ?? 0;
        }
        catch
        {
            return 0;
        }
    }

    private static CandidateBuilder GetOrCreateCandidate(
        IDictionary<string, CandidateBuilder> candidateMap,
        string packedProductCode)
    {
        if (candidateMap.TryGetValue(packedProductCode, out var existing))
        {
            return existing;
        }

        var created = new CandidateBuilder(packedProductCode);
        candidateMap[packedProductCode] = created;
        return created;
    }

    private static IReadOnlyCollection<RunningMsiexecInfo> DetectRunningMsiexec()
    {
        var processes = new List<RunningMsiexecInfo>();

        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT ProcessId, CommandLine FROM Win32_Process WHERE Name = 'msiexec.exe'");

            foreach (ManagementObject process in searcher.Get())
            {
                var processId = Convert.ToInt32(process["ProcessId"]);
                var commandLine = Convert.ToString(process["CommandLine"]) ?? string.Empty;
                var isAutodeskRelated = MatchesAutodesk(commandLine);

                processes.Add(new RunningMsiexecInfo(processId, commandLine, isAutodeskRelated));
            }
        }
        catch (Exception ex)
        {
            Logger.Debug(ex, "Failed to query running msiexec processes");
        }

        return processes
            .OrderBy(process => process.ProcessId)
            .ToList();
    }

    private static bool MatchesAutodesk(string? value)
    {
        return !string.IsNullOrWhiteSpace(value) && AutodeskPattern.IsMatch(value);
    }

    private static string? ReadRegistryString(RegistryKey? key, string valueName)
    {
        return key?.GetValue(valueName) as string;
    }

    private static string? FirstNonEmpty(params string?[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
    }

    private static bool HasOnlyMediaEntries(RegistryKey? sourceListKey, RegistryKey? mediaKey, string? packageName, string? lastUsedSource)
    {
        if (sourceListKey is null || mediaKey is null)
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(packageName) || !string.IsNullOrWhiteSpace(lastUsedSource))
        {
            return false;
        }

        return mediaKey.ValueCount > 0 || mediaKey.GetValueNames().Length > 0 || mediaKey.GetSubKeyNames().Length > 0;
    }

    private static string ExtractSourcePath(string lastUsedSource)
    {
        if (string.IsNullOrWhiteSpace(lastUsedSource))
        {
            return string.Empty;
        }

        var parts = lastUsedSource.Split(';', StringSplitOptions.RemoveEmptyEntries);
        return parts.Length == 0 ? lastUsedSource : parts[^1].Trim();
    }

    private static bool IsTemporaryPath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        return path.Contains(@"\Temp\", StringComparison.OrdinalIgnoreCase) ||
               path.Contains("%TEMP%", StringComparison.OrdinalIgnoreCase);
    }

    private static bool PathExists(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return false;
        }

        return Directory.Exists(path) || File.Exists(path);
    }

    private static string GetCandidateLabel(AutodeskMsiRegistrationCandidate candidate)
    {
        if (!string.IsNullOrWhiteSpace(candidate.ProductLine) && !string.IsNullOrWhiteSpace(candidate.DisplayName))
        {
            return $"{candidate.ProductLine} - {candidate.DisplayName}";
        }

        if (!string.IsNullOrWhiteSpace(candidate.DisplayName))
        {
            return candidate.DisplayName;
        }

        if (!string.IsNullOrWhiteSpace(candidate.PackageName))
        {
            return candidate.PackageName;
        }

        return candidate.PackedProductCode;
    }

    private sealed class CandidateBuilder
    {
        public CandidateBuilder(string packedProductCode)
        {
            PackedProductCode = packedProductCode;
        }

        public string PackedProductCode { get; }
        public string DisplayName { get; set; } = string.Empty;
        public string? Publisher { get; set; }
        public string? ProductCode { get; set; }
        public string? PackageName { get; set; }
        public string? PackageDisplayName { get; set; }
        public string? LastUsedSource { get; set; }
        public string? MediaPath { get; set; }
        public string? ProductLine { get; set; }
        public bool IsAutodeskRelated { get; set; }
        public bool IsStale { get; set; }
        public HashSet<string> Reasons { get; } = new(StringComparer.OrdinalIgnoreCase);

        public AutodeskMsiRegistrationCandidate Build()
        {
            var displayName = string.IsNullOrWhiteSpace(PackageDisplayName)
                ? (string.IsNullOrWhiteSpace(DisplayName) ? PackedProductCode : DisplayName)
                : PackageDisplayName;

            return new AutodeskMsiRegistrationCandidate(
                PackedProductCode: PackedProductCode,
                ProductCode: ProductCode,
                DisplayName: displayName,
                Publisher: Publisher,
                PackageName: PackageName,
                LastUsedSource: LastUsedSource,
                MediaPath: MediaPath,
                ProductLine: ProductLine,
                Reasons: Reasons.OrderBy(reason => reason, StringComparer.OrdinalIgnoreCase).ToList());
        }
    }
}

/// <summary>
/// A Windows Installer product registration selected for Autodesk cleanup.
/// </summary>
/// <param name="PackedProductCode">The packed Installer product code.</param>
/// <param name="ProductCode">The original MSI product code when known.</param>
/// <param name="DisplayName">The product display name.</param>
/// <param name="Publisher">The product publisher.</param>
/// <param name="PackageName">The recorded package name.</param>
/// <param name="LastUsedSource">The recorded last used source.</param>
/// <param name="MediaPath">The installer media path when identified.</param>
/// <param name="ProductLine">The Autodesk product line when identified.</param>
/// <param name="Reasons">The reasons this registration was selected.</param>
internal readonly record struct AutodeskMsiRegistrationCandidate(
    string PackedProductCode,
    string? ProductCode,
    string DisplayName,
    string? Publisher,
    string? PackageName,
    string? LastUsedSource,
    string? MediaPath,
    string? ProductLine,
    IReadOnlyCollection<string> Reasons)
{
    public IReadOnlyCollection<string> RegistryPaths =>
    [
        $@"SOFTWARE\Classes\Installer\Products\{PackedProductCode}",
        $@"SOFTWARE\Classes\Installer\Features\{PackedProductCode}",
        $@"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\{PackedProductCode}"
    ];
}

/// <summary>
/// A running msiexec process discovered during cleanup.
/// </summary>
/// <param name="ProcessId">The msiexec PID.</param>
/// <param name="CommandLine">The process command line.</param>
/// <param name="IsAutodeskRelated">Whether the command line appears Autodesk-related.</param>
internal readonly record struct RunningMsiexecInfo(int ProcessId, string CommandLine, bool IsAutodeskRelated);

/// <summary>
/// The result of Autodesk MSI registration cleanup.
/// </summary>
/// <param name="Candidates">The MSI registrations selected for cleanup.</param>
/// <param name="ParsedFailures">The parsed failure records.</param>
/// <param name="RunningMsiexec">Detected running msiexec processes.</param>
/// <param name="SuccessfulRemovals">The number of candidates removed or planned.</param>
/// <param name="Errors">Any errors encountered.</param>
internal readonly record struct BrokenMsiCleanupResult(
    IReadOnlyCollection<AutodeskMsiRegistrationCandidate> Candidates,
    IReadOnlyCollection<RunningMsiexecInfo> RunningMsiexec,
    int SuccessfulRemovals,
    IReadOnlyCollection<string> Errors)
{
    public bool IsSuccessful => Errors.Count == 0;
}
