namespace OpenXmlKit.Word;

/// <summary>
/// The formatting a table style applies to one part of a table — the header row, the banding, a
/// corner cell.
/// </summary>
/// <remarks>
/// This is what makes a table style a table style rather than a set of borders: <c>TableGrid</c>
/// and every banded style Word ships are a whole-table block plus a handful of these. Which of
/// them a given table honours is the table's own <see cref="TableFormat.Look"/>.
/// <para>
/// The schema narrows each block: a conditional override carries no style reference, no table
/// width and no cell span, because those belong to the content rather than to the style. Stating
/// one here throws rather than being dropped on the way out, since a dropped child is a style that
/// silently does less than it says.
/// </para>
/// </remarks>
public class TableStyleConditional :
    ITableStyleConditionalView
{
    internal TableStyleConditional(TableStyleArea area) =>
        Area = area;

    /// <summary>
    /// The part of the table this block formats.
    /// </summary>
    public TableStyleArea Area { get; }

    /// <summary>
    /// Character formatting for text in this part of the table.
    /// </summary>
    public Font Font { get; } = new();

    /// <summary>
    /// Paragraph formatting for text in this part of the table.
    /// </summary>
    public ParagraphFormat ParagraphFormat { get; } = new();

    /// <summary>
    /// Table-level formatting — borders, cell margins, alignment.
    /// </summary>
    public TableFormat TableFormat { get; } = new();

    /// <summary>
    /// Cell-level formatting — borders, shading, margins, vertical alignment.
    /// </summary>
    public CellFormat CellFormat { get; } = new();

    IFontView ITableStyleConditionalView.Font => Font;
    IParagraphFormatView ITableStyleConditionalView.ParagraphFormat => ParagraphFormat;
    ITableFormatView ITableStyleConditionalView.TableFormat => TableFormat;
    ICellFormatView ITableStyleConditionalView.CellFormat => CellFormat;

    /// <summary>
    /// Whether the block states any formatting. An empty block is not written.
    /// </summary>
    public bool IsEmpty =>
        Font.IsEmpty &&
        ParagraphFormat.IsEmpty &&
        TableFormat.IsEmpty &&
        CellFormat.IsEmpty;

    /// <summary>
    /// Overwrites every part with the other block.
    /// </summary>
    public void CopyFrom(TableStyleConditional other)
    {
        Font.CopyFrom(other.Font);
        ParagraphFormat.CopyFrom(other.ParagraphFormat);
        TableFormat.CopyFrom(other.TableFormat);
        CellFormat.CopyFrom(other.CellFormat);
    }

    internal W.TableStyleProperties ToOpenXml()
    {
        var element = new W.TableStyleProperties
        {
            Type = Area.ToOpenXml()
        };

        if (!ParagraphFormat.IsEmpty)
        {
            element.StyleParagraphProperties =
                Transfer<W.StyleParagraphProperties>(ParagraphFormat.ToProperties(), AllowedInParagraph, "paragraph");
        }

        if (!Font.IsEmpty)
        {
            element.RunPropertiesBaseStyle =
                Transfer<W.RunPropertiesBaseStyle>(Font.ToProperties(), AllowedInRun, "character");
        }

        if (!TableFormat.IsEmpty)
        {
            element.TableStyleConditionalFormattingTableProperties =
                Transfer<W.TableStyleConditionalFormattingTableProperties>(TableFormat.ToProperties(), AllowedInTable, "table");
        }

        if (!CellFormat.IsEmpty)
        {
            element.TableStyleConditionalFormattingTableCellProperties =
                Transfer<W.TableStyleConditionalFormattingTableCellProperties>(CellFormat.ToProperties(), AllowedInCell, "cell");
        }

        return element;
    }

    // The four allow-lists below are stated as the children the *_StyleOverride schemas exclude,
    // rather than as the much longer lists they permit. Each exclusion is a property that only
    // means something on a piece of content: a style cannot name another style, set its own width,
    // or span columns.

    static bool AllowedInRun(OpenXmlElement child) =>
        child is not (W.RunStyle or W.Highlight or W.RightToLeftText);

    static bool AllowedInParagraph(OpenXmlElement child) =>
        child is not W.ParagraphStyleId;

    static bool AllowedInTable(OpenXmlElement child) =>
        child is not (W.TableStyle or W.TableWidth or W.TableLayout or W.TableLook or W.TableCaption or W.TableDescription);

    static bool AllowedInCell(OpenXmlElement child) =>
        child is not (W.TableCellWidth or W.GridSpan or W.VerticalMerge or W.TextDirection or W.TableCellFitText);

    static T Transfer<T>(OpenXmlElement? source, Func<OpenXmlElement, bool> allowed, string what)
        where T : OpenXmlElement, new()
    {
        var target = new T();
        if (source == null)
        {
            return target;
        }

        List<string>? rejected = null;
        foreach (var child in source.ChildElements.ToList())
        {
            child.Remove();
            if (allowed(child))
            {
                target.AppendChild(child);
                continue;
            }

            rejected ??= [];
            rejected.Add(child.LocalName);
        }

        if (rejected != null)
        {
            throw new InvalidOperationException(
                $"A table style's conditional {what} formatting cannot state {string.Join(", ", rejected)}. " +
                "Those belong on the content that uses the style rather than on the style itself.");
        }

        return target;
    }
}
