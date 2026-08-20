/// <summary>
/// Compiles only if the alias props works: the SDK's names and this library's are both in scope
/// under names that do not collide.
/// </summary>
/// <remarks>
/// A compile-time check with no test to run. If the two ever do collide the failure is CS0104 at
/// build time, which is the failure a consumer would otherwise be the first to see.
/// </remarks>
static class BothInOneFile
{
    /// <summary>
    /// The SDK's Paragraph, unqualified and unambiguous.
    /// </summary>
    public static Paragraph RawParagraph() =>
        new(new Run(new Text("built with the SDK")));

    /// <summary>
    /// This library's, under its alias, alongside the one above.
    /// </summary>
    public static Paragraph WrappedParagraph()
    {
        var paragraph = new WParagraph("built with OpenXmlKit")
        {
            Format =
            {
                Alignment = WParagraphAlignment.Center
            }
        };
        paragraph.AddRun("and bold").Font.Bold = WToggle.On;
        return paragraph.ToOpenXml();
    }

    /// <summary>
    /// A table, both ways, to cover the names that collide most.
    /// </summary>
    public static Table WrappedTable() =>
        WTable.Create()
            .Width(WWidth.Percent(100))
            .Borders(WBorderStyle.Single, WLength.FromPoints(0.5), WColor.Black)
            .Row("a", "b")
            .ToOpenXml();

    /// <summary>
    /// The escape hatch in both directions: a raw element into the wrapper, and back out.
    /// </summary>
    public static byte[] Mixed()
    {
        using var document = WDocument.Create();
        document.Body.AppendElement(RawParagraph());
        document.Body.AppendElement(WrappedParagraph());
        document.Body.AppendElement(WrappedTable());

        var styles = document.Styles;
        styles.EnsureBuiltIn(WBuiltInStyle.TableGrid);
        return document.ToArray();
    }
}

/// <summary>
/// The read half, under aliases, in the same file as the SDK's own names.
/// </summary>
static class ReadingUnderAliases
{
    /// <summary>
    /// Reading gives back views, and a view has no way to change what it is looking at. The lines
    /// commented out below are the point: they do not compile, which is the guarantee.
    /// </summary>
    public static string FirstParagraph(byte[] bytes)
    {
        using var document = WDocumentView.Open(bytes);
        var paragraph = document.Body.Paragraphs.First();

        // paragraph.AddBookmark(...);              no such method on a view
        // paragraph.Format.Alignment = ...;        IWParagraphFormatView has no setters

        // Typed as IWFontView, the alias for IFontView: the prefix goes after the leading I.
        var font = document.Formatting.FontFor(paragraph.Runs.First(), paragraph);
        return $"{paragraph.Text} in {font.Name ?? "the default font"}";
    }

    /// <summary>
    /// And the SDK's own reading, unqualified, alongside it.
    /// </summary>
    public static int CountParagraphs(MainDocumentPart main) =>
        main.Document!.Body!.Elements<Paragraph>().Count();
}
