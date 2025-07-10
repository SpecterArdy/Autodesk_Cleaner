using System.Runtime.InteropServices;

namespace Autodesk_Cleaner.Core;

/// <summary>
/// P/Invoke declarations for interacting with Windows registry APIs.
/// </summary>
internal static class NativeMethods
{
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegOpenKeyEx(
        UIntPtr hKey,
        string subKey,
        uint ulOptions,
        int samDesired,
        out UIntPtr phkResult);

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegCloseKey(
        UIntPtr hKey);

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegDeleteValue(
        UIntPtr hKey,
        string lpValueName);

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern int RegDeleteKeyEx(
        UIntPtr hKey,
        string lpSubKey,
        int samDesired,
        uint Reserved);

    public const int KEY_READ = 0x20019;
    public const int KEY_WRITE = 0x20006;
    public const int KEY_WOW64_64KEY = 0x0100;
    public const int REG_OPTION_OPEN_LINK = 0x00000008;
}

