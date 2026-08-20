namespace OpenXmlKit.Word;

/// <summary>
/// One part of a table style's conditional formatting, as read.
/// </summary>
public interface ITableStyleConditionalView
{
    TableStyleArea Area { get; }
    IFontView Font { get; }
    IParagraphFormatView ParagraphFormat { get; }
    ITableFormatView TableFormat { get; }
    ICellFormatView CellFormat { get; }
    bool IsEmpty { get; }
}
