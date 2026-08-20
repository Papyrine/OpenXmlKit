namespace OpenXmlKit.Word;

/// <summary>
/// One part of a table style's conditional formatting, as read.
/// </summary>
public interface ITableStyleConditionalView
{
    /// <inheritdoc cref="TableStyleConditional.Area"/>
    TableStyleArea Area { get; }

    /// <inheritdoc cref="TableStyleConditional.Font"/>
    IFontView Font { get; }

    /// <inheritdoc cref="TableStyleConditional.ParagraphFormat"/>
    IParagraphFormatView ParagraphFormat { get; }

    /// <inheritdoc cref="TableStyleConditional.TableFormat"/>
    ITableFormatView TableFormat { get; }

    /// <inheritdoc cref="TableStyleConditional.CellFormat"/>
    ICellFormatView CellFormat { get; }

    /// <inheritdoc cref="TableStyleConditional.IsEmpty"/>
    bool IsEmpty { get; }
}
