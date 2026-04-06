namespace Autodesk_Cleaner.Core;

/// <summary>
/// Converts MSI product codes into Windows Installer packed registry keys.
/// </summary>
internal static class MsiPackedGuid
{
    /// <summary>
    /// Converts a GUID to the packed Windows Installer product code form.
    /// </summary>
    /// <param name="guid">The product GUID.</param>
    /// <returns>The packed product code used under Installer registry hives.</returns>
    public static string ToPackedProductCode(Guid guid)
    {
        var value = guid.ToString("N").ToUpperInvariant();

        var part1 = ReversePairs(value[..8]);
        var part2 = ReversePairs(value.Substring(8, 4));
        var part3 = ReversePairs(value.Substring(12, 4));
        var part4 = ReverseNibblesPerByte(value.Substring(16, 4));
        var part5 = ReverseNibblesPerByte(value.Substring(20, 12));

        return part1 + part2 + part3 + part4 + part5;
    }

    /// <summary>
    /// Attempts to convert a product code string to packed form.
    /// </summary>
    /// <param name="productCode">The MSI product code.</param>
    /// <param name="packedProductCode">The packed code when conversion succeeds.</param>
    /// <returns>True when the product code was valid.</returns>
    public static bool TryToPackedProductCode(string? productCode, out string packedProductCode)
    {
        packedProductCode = string.Empty;
        if (!Guid.TryParse(productCode, out var guid))
        {
            return false;
        }

        packedProductCode = ToPackedProductCode(guid);
        return true;
    }

    private static string ReversePairs(string value)
    {
        return string.Concat(Enumerable
            .Range(0, value.Length / 2)
            .Select(index => value.Substring(index * 2, 2))
            .Reverse());
    }

    private static string ReverseNibblesPerByte(string value)
    {
        return string.Concat(Enumerable
            .Range(0, value.Length / 2)
            .Select(index => value.Substring(index * 2, 2))
            .Select(pair => $"{pair[1]}{pair[0]}"));
    }
}
