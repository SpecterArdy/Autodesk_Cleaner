using Microsoft.Win32;
using NLog;
using System.Diagnostics;
using System.ServiceProcess;

namespace Autodesk_Cleaner.Core;

/// <summary>
/// Removes Autodesk licensing contamination that can interfere with cleanup or reinstall.
/// </summary>
internal sealed class AutodeskLicensingCleaner
{
    private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
    private readonly ScannerConfig _config;

    public AutodeskLicensingCleaner(ScannerConfig config)
    {
        _config = config;
    }

    /// <summary>
    /// Removes stale Autodesk network-license overrides and related local service state.
    /// </summary>
    /// <param name="includeHelperReset">Whether to invoke AdskLicensingInstHelper.exe.</param>
    /// <param name="productKey">Optional product key for helper reset.</param>
    /// <param name="productVersion">Optional product version for helper reset.</param>
    /// <returns>A licensing cleanup result.</returns>
    public async Task<AutodeskLicensingCleanupResult> CleanupAsync(
        bool includeHelperReset = false,
        string? productKey = null,
        string? productVersion = null)
    {
        var steps = new List<LicensingCleanupAction>();
        var errors = new List<string>();

        if (includeHelperReset && !string.IsNullOrWhiteSpace(productKey) && !string.IsNullOrWhiteSpace(productVersion))
        {
            var helperStep = await ResetLicenseModeAsync(productKey, productVersion);
            steps.Add(helperStep);
            if (helperStep.Status.Equals("Failed", StringComparison.OrdinalIgnoreCase))
            {
                errors.Add(helperStep.Details);
            }
        }

        foreach (var target in GetFlexLmTargets())
        {
            var step = RemoveRegistryValueIfPresent(target.Hive, target.View, target.KeyPath, "ADSKFLEX_LICENSE_FILE", target.DisplayPath);
            steps.Add(step);
            if (step.Status.Equals("Failed", StringComparison.OrdinalIgnoreCase))
            {
                errors.Add(step.Details);
            }
        }

        var adskNlmStep = await SetServiceStartupManualAsync("AdskNLM");
        steps.Add(adskNlmStep);
        if (adskNlmStep.Status.Equals("Failed", StringComparison.OrdinalIgnoreCase))
        {
            errors.Add(adskNlmStep.Details);
        }

        Logger.Info("Autodesk licensing cleanup completed: {@Steps}", steps);

        return new AutodeskLicensingCleanupResult(
            Steps: steps,
            Errors: errors);
    }

    private LicensingCleanupAction RemoveRegistryValueIfPresent(
        RegistryHive hive,
        RegistryView view,
        string keyPath,
        string valueName,
        string displayPath)
    {
        try
        {
            using var baseKey = RegistryKey.OpenBaseKey(hive, view);
            using var subKey = baseKey.OpenSubKey(keyPath, writable: !_config.DryRun);

            if (subKey is null)
            {
                return new LicensingCleanupAction("Remove stale FLEXlm override", "Not present", displayPath);
            }

            var existingValue = subKey.GetValue(valueName);
            if (existingValue is null)
            {
                return new LicensingCleanupAction("Remove stale FLEXlm override", "Not present", displayPath);
            }

            var existingValueText = Convert.ToString(existingValue) ?? "<non-string>";

            if (_config.DryRun)
            {
                return new LicensingCleanupAction(
                    "Remove stale FLEXlm override",
                    "Planned",
                    $"{displayPath} = '{existingValueText}'");
            }

            subKey.DeleteValue(valueName, throwOnMissingValue: false);
            return new LicensingCleanupAction(
                "Remove stale FLEXlm override",
                "Removed",
                $"{displayPath} = '{existingValueText}'");
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Failed to remove FLEXlm value at {DisplayPath}", displayPath);
            return new LicensingCleanupAction("Remove stale FLEXlm override", "Failed", $"{displayPath}: {ex.Message}");
        }
    }

    private async Task<LicensingCleanupAction> SetServiceStartupManualAsync(string serviceName)
    {
        try
        {
            var isInstalled = ServiceController.GetServices()
                .Any(service => string.Equals(service.ServiceName, serviceName, StringComparison.OrdinalIgnoreCase));

            if (!isInstalled)
            {
                return new LicensingCleanupAction(
                    "Set AdskNLM startup to Manual",
                    "Not installed",
                    $"{serviceName} is not registered.");
            }

            if (_config.DryRun)
            {
                return new LicensingCleanupAction(
                    "Set AdskNLM startup to Manual",
                    "Planned",
                    $"{serviceName} startup will be set to Manual.");
            }

            var (success, standardOutput, standardError) =
                await RunProcessAsync("sc.exe", $"config \"{serviceName}\" start= demand");

            var details = success
                ? $"{serviceName} startup set to Manual."
                : string.IsNullOrWhiteSpace(standardError) ? standardOutput.Trim() : standardError.Trim();

            return new LicensingCleanupAction(
                "Set AdskNLM startup to Manual",
                success ? "Applied" : "Failed",
                details);
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "Failed to set {ServiceName} startup to Manual", serviceName);
            return new LicensingCleanupAction("Set AdskNLM startup to Manual", "Failed", ex.Message);
        }
    }

    private async Task<LicensingCleanupAction> ResetLicenseModeAsync(string productKey, string productVersion)
    {
        var helperPath = GetLicensingHelperPath();
        if (!File.Exists(helperPath))
        {
            return new LicensingCleanupAction(
                "Reset named-user licensing",
                "Skipped",
                $"Licensing helper not found at {helperPath}");
        }

        if (_config.DryRun)
        {
            return new LicensingCleanupAction(
                "Reset named-user licensing",
                "Planned",
                $"Would run helper for product {productKey} {productVersion}");
        }

        var arguments = $"change -pk {productKey} -pv {productVersion} -lm \"\"";
        var (success, standardOutput, standardError) = await RunProcessAsync(helperPath, arguments);
        var details = success
            ? string.IsNullOrWhiteSpace(standardOutput) ? "Licensing helper completed." : standardOutput.Trim()
            : string.IsNullOrWhiteSpace(standardError) ? "Licensing helper returned a non-zero exit code." : standardError.Trim();

        return new LicensingCleanupAction(
            "Reset named-user licensing",
            success ? "Applied" : "Failed",
            details);
    }

    private static IEnumerable<(RegistryHive Hive, RegistryView View, string KeyPath, string DisplayPath)> GetFlexLmTargets()
    {
        yield return (RegistryHive.CurrentUser, RegistryView.Default, @"Software\FLEXlm License Manager", @"HKCU\Software\FLEXlm License Manager\ADSKFLEX_LICENSE_FILE");
        yield return (RegistryHive.LocalMachine, RegistryView.Registry64, @"Software\FLEXlm License Manager", @"HKLM\Software\FLEXlm License Manager\ADSKFLEX_LICENSE_FILE");
        yield return (RegistryHive.LocalMachine, RegistryView.Registry32, @"Software\FLEXlm License Manager", @"HKLM\Software\Wow6432Node\FLEXlm License Manager\ADSKFLEX_LICENSE_FILE");
    }

    private static string GetLicensingHelperPath()
    {
        var commonProgramFilesX86 = Environment.GetEnvironmentVariable("CommonProgramFiles(x86)");
        if (string.IsNullOrWhiteSpace(commonProgramFilesX86))
        {
            commonProgramFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86);
        }

        return Path.Combine(
            commonProgramFilesX86,
            "Autodesk Shared",
            "AdskLicensing",
            "Current",
            "helper",
            "AdskLicensingInstHelper.exe");
    }

    private static async Task<(bool Success, string StandardOutput, string StandardError)> RunProcessAsync(string fileName, string arguments)
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
            return (false, string.Empty, $"Failed to start process: {fileName}");
        }

        var standardOutputTask = process.StandardOutput.ReadToEndAsync();
        var standardErrorTask = process.StandardError.ReadToEndAsync();

        await process.WaitForExitAsync();
        return (process.ExitCode == 0, await standardOutputTask, await standardErrorTask);
    }
}

/// <summary>
/// A single licensing cleanup action.
/// </summary>
/// <param name="Step">The action name.</param>
/// <param name="Status">The action status.</param>
/// <param name="Details">The operator-facing detail text.</param>
internal readonly record struct LicensingCleanupAction(string Step, string Status, string Details);

/// <summary>
/// The result of Autodesk licensing cleanup.
/// </summary>
/// <param name="Steps">The actions executed or planned.</param>
/// <param name="Errors">Any failures encountered.</param>
internal readonly record struct AutodeskLicensingCleanupResult(
    IReadOnlyCollection<LicensingCleanupAction> Steps,
    IReadOnlyCollection<string> Errors)
{
    public bool IsSuccessful => Errors.Count == 0;
}
