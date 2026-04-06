using Microsoft.Win32;
using NLog;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;

namespace Autodesk_Cleaner.Core;

/// <summary>
/// Handles privilege elevation and ACL takeover for protected registry keys.
/// </summary>
internal static class RegistryAclHelper
{
    private static readonly string[] RequiredPrivileges =
    [
        "SeTakeOwnershipPrivilege",
        "SeRestorePrivilege",
        "SeBackupPrivilege"
    ];

    /// <summary>
    /// Enables the privileges required for ownership and backup operations.
    /// </summary>
    /// <param name="logger">Optional logger.</param>
    /// <returns>True when all requested privileges were enabled.</returns>
    public static bool EnableRequiredPrivileges(Logger? logger = null)
    {
        var allEnabled = true;

        foreach (var privilege in RequiredPrivileges)
        {
            if (!EnablePrivilege(privilege, logger))
            {
                allEnabled = false;
            }
        }

        return allEnabled;
    }

    /// <summary>
    /// Exports a registry key to a .reg file when it exists.
    /// </summary>
    /// <param name="fullRegistryPath">The full reg.exe path, for example HKLM\SOFTWARE\Foo.</param>
    /// <param name="backupFilePath">The destination .reg file path.</param>
    /// <param name="logger">Optional logger.</param>
    /// <returns>True when the export succeeds or the key does not exist.</returns>
    public static async Task<bool> ExportKeyAsync(string fullRegistryPath, string backupFilePath, Logger? logger = null)
    {
        try
        {
            var directoryPath = Path.GetDirectoryName(backupFilePath);
            if (!string.IsNullOrWhiteSpace(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = "reg.exe",
                Arguments = $"export \"{fullRegistryPath}\" \"{backupFilePath}\" /y",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            using var process = Process.Start(startInfo);
            if (process is null)
            {
                return false;
            }

            await process.WaitForExitAsync();
            if (process.ExitCode == 0)
            {
                return true;
            }

            var standardError = await process.StandardError.ReadToEndAsync();
            if (standardError.Contains("unable to find", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            logger?.Debug("Registry export failed for {RegistryPath}: {Error}", fullRegistryPath, standardError.Trim());
            return false;
        }
        catch (Exception ex)
        {
            logger?.Debug(ex, "Registry export failed for {RegistryPath}", fullRegistryPath);
            return false;
        }
    }

    /// <summary>
    /// Grants Administrators full control recursively on a key tree.
    /// </summary>
    /// <param name="hive">The target hive.</param>
    /// <param name="view">The registry view.</param>
    /// <param name="keyPath">The key path relative to the hive.</param>
    /// <param name="logger">Optional logger.</param>
    /// <returns>True when the ACL update succeeds.</returns>
    public static bool GrantFullControlRecursive(RegistryHive hive, RegistryView view, string keyPath, Logger? logger = null)
    {
        try
        {
            EnableRequiredPrivileges(logger);

            using var baseKey = RegistryKey.OpenBaseKey(hive, view);
            var administratorsSid = new SecurityIdentifier(WellKnownSidType.BuiltinAdministratorsSid, null);

            return GrantFullControlRecursive(baseKey, keyPath, administratorsSid, logger);
        }
        catch (Exception ex)
        {
            logger?.Debug(ex, "Failed to grant recursive registry access on {Hive}\\{KeyPath}", hive, keyPath);
            return false;
        }
    }

    /// <summary>
    /// Deletes a registry key tree, taking ownership first if required.
    /// </summary>
    /// <param name="hive">The target hive.</param>
    /// <param name="view">The registry view.</param>
    /// <param name="keyPath">The key path relative to the hive.</param>
    /// <param name="logger">Optional logger.</param>
    /// <returns>True when the key no longer exists.</returns>
    public static bool DeleteTree(RegistryHive hive, RegistryView view, string keyPath, Logger? logger = null)
    {
        try
        {
            using var baseKey = RegistryKey.OpenBaseKey(hive, view);
            if (!KeyExists(baseKey, keyPath))
            {
                return true;
            }

            var parentPath = GetParentPath(keyPath);
            var leafName = GetLeafName(keyPath);
            if (string.IsNullOrWhiteSpace(parentPath) || string.IsNullOrWhiteSpace(leafName))
            {
                return false;
            }

            if (TryDelete(baseKey, parentPath, leafName))
            {
                return true;
            }

            GrantFullControlRecursive(hive, view, keyPath, logger);
            GrantFullControlRecursive(hive, view, parentPath, logger);

            return TryDelete(baseKey, parentPath, leafName) || !KeyExists(baseKey, keyPath);
        }
        catch (Exception ex)
        {
            logger?.Debug(ex, "Failed to delete registry tree {Hive}\\{KeyPath}", hive, keyPath);
            return false;
        }
    }

    /// <summary>
    /// Checks whether a registry key exists.
    /// </summary>
    /// <param name="hive">The hive.</param>
    /// <param name="view">The view.</param>
    /// <param name="keyPath">The relative key path.</param>
    /// <returns>True when the key exists.</returns>
    public static bool KeyExists(RegistryHive hive, RegistryView view, string keyPath)
    {
        using var baseKey = RegistryKey.OpenBaseKey(hive, view);
        return KeyExists(baseKey, keyPath);
    }

    private static bool EnablePrivilege(string privilegeName, Logger? logger)
    {
        if (!NativeMethods.OpenProcessToken(
                Process.GetCurrentProcess().Handle,
                NativeMethods.TOKEN_ADJUST_PRIVILEGES | NativeMethods.TOKEN_QUERY,
                out var tokenHandle))
        {
            logger?.Debug("OpenProcessToken failed while enabling {PrivilegeName}: {Error}",
                privilegeName,
                new Win32Exception(Marshal.GetLastWin32Error()).Message);
            return false;
        }

        try
        {
            if (!NativeMethods.LookupPrivilegeValue(null, privilegeName, out var luid))
            {
                logger?.Debug("LookupPrivilegeValue failed for {PrivilegeName}: {Error}",
                    privilegeName,
                    new Win32Exception(Marshal.GetLastWin32Error()).Message);
                return false;
            }

            var tokenPrivileges = new NativeMethods.TokenPrivileges
            {
                PrivilegeCount = 1,
                Luid = luid,
                Attributes = NativeMethods.SE_PRIVILEGE_ENABLED
            };

            if (!NativeMethods.AdjustTokenPrivileges(tokenHandle, false, ref tokenPrivileges, 0, nint.Zero, nint.Zero))
            {
                logger?.Debug("AdjustTokenPrivileges failed for {PrivilegeName}: {Error}",
                    privilegeName,
                    new Win32Exception(Marshal.GetLastWin32Error()).Message);
                return false;
            }

            return Marshal.GetLastWin32Error() == 0;
        }
        finally
        {
            NativeMethods.CloseHandle(tokenHandle);
        }
    }

    private static bool GrantFullControlRecursive(
        RegistryKey baseKey,
        string keyPath,
        SecurityIdentifier administratorsSid,
        Logger? logger)
    {
        try
        {
            using (var ownershipKey = baseKey.OpenSubKey(
                       keyPath,
                       RegistryKeyPermissionCheck.ReadWriteSubTree,
                       RegistryRights.TakeOwnership | RegistryRights.ReadKey | RegistryRights.ChangePermissions))
            {
                if (ownershipKey is null)
                {
                    return false;
                }

                var ownerSecurity = ownershipKey.GetAccessControl(AccessControlSections.Owner);
                ownerSecurity.SetOwner(administratorsSid);
                ownershipKey.SetAccessControl(ownerSecurity);
            }

            using var permissionKey = baseKey.OpenSubKey(
                keyPath,
                RegistryKeyPermissionCheck.ReadWriteSubTree,
                RegistryRights.ReadKey | RegistryRights.ChangePermissions | RegistryRights.FullControl);

            if (permissionKey is null)
            {
                return false;
            }

            var security = permissionKey.GetAccessControl();
            security.SetOwner(administratorsSid);
            security.AddAccessRule(new RegistryAccessRule(
                administratorsSid,
                RegistryRights.FullControl,
                InheritanceFlags.ContainerInherit,
                PropagationFlags.None,
                AccessControlType.Allow));
            permissionKey.SetAccessControl(security);

            foreach (var subKeyName in permissionKey.GetSubKeyNames())
            {
                GrantFullControlRecursive(baseKey, $@"{keyPath}\{subKeyName}", administratorsSid, logger);
            }

            return true;
        }
        catch (Exception ex)
        {
            logger?.Debug(ex, "Failed to grant recursive registry access on {KeyPath}", keyPath);
            return false;
        }
    }

    private static bool KeyExists(RegistryKey baseKey, string keyPath)
    {
        using var key = baseKey.OpenSubKey(keyPath, writable: false);
        return key is not null;
    }

    private static bool TryDelete(RegistryKey baseKey, string parentPath, string leafName)
    {
        using var parentKey = baseKey.OpenSubKey(parentPath, writable: true);
        if (parentKey is null)
        {
            return true;
        }

        if (!parentKey.GetSubKeyNames().Contains(leafName, StringComparer.OrdinalIgnoreCase))
        {
            return true;
        }

        parentKey.DeleteSubKeyTree(leafName, throwOnMissingSubKey: false);
        return !parentKey.GetSubKeyNames().Contains(leafName, StringComparer.OrdinalIgnoreCase);
    }

    private static string? GetParentPath(string keyPath)
    {
        var separatorIndex = keyPath.LastIndexOf('\\');
        return separatorIndex > 0 ? keyPath[..separatorIndex] : null;
    }

    private static string? GetLeafName(string keyPath)
    {
        var separatorIndex = keyPath.LastIndexOf('\\');
        return separatorIndex >= 0 && separatorIndex < keyPath.Length - 1
            ? keyPath[(separatorIndex + 1)..]
            : null;
    }
}
