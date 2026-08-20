using DocumentFormat.OpenXml.Validation;

/// <summary>
/// Asserts that a produced document is one Word will open without offering to repair it.
/// </summary>
/// <remarks>
/// This is the check the emitter design exists to pass. A schema violation — most often a
/// properties element whose children are out of their CT_* sequence — does not throw when it is
/// written; it surfaces later as Word declaring the file corrupt and silently stripping the
/// formatting on repair. Running the validator over everything the tests produce is what turns
/// that into a failing build instead.
/// </remarks>
public static class DocumentAssert
{
    public static byte[] IsValid(Document document)
    {
        var bytes = document.ToArray();
        IsValid(bytes);
        return bytes;
    }

    public static void IsValid(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var package = WordprocessingDocument.Open(stream, false);

        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        var errors = validator.Validate(package).ToList();
        if (errors.Count == 0)
        {
            return;
        }

        var report = new StringBuilder($"{errors.Count} schema validation error(s):");
        foreach (var error in errors)
        {
            report.AppendLine();
            report.AppendLine($"  {error.Description}");
            report.AppendLine($"    part: {error.Part?.Uri}");
            report.AppendLine($"    path: {error.Path?.XPath}");
        }

        Assert.Fail(report.ToString());
    }

    /// <summary>
    /// The XML of the main document part, for asserting on what was actually written.
    /// </summary>
    public static string MainPartXml(Document document)
    {
        using var stream = new MemoryStream(document.ToArray());
        using var package = WordprocessingDocument.Open(stream, false);
        return package.MainDocumentPart!.Document!.OuterXml;
    }

    /// <summary>
    /// The XML of a named part, or null when the document has none.
    /// </summary>
    public static string? PartXml(Document document, string partName)
    {
        using var stream = new MemoryStream(document.ToArray());
        using var package = WordprocessingDocument.Open(stream, false);
        var main = package.MainDocumentPart!;
        return partName switch
        {
            "styles" => main.StyleDefinitionsPart?.Styles?.OuterXml,
            "numbering" => main.NumberingDefinitionsPart?.Numbering?.OuterXml,
            "settings" => main.DocumentSettingsPart?.Settings?.OuterXml,
            _ => throw new ArgumentOutOfRangeException(nameof(partName), partName, "Unknown part.")
        };
    }
}
