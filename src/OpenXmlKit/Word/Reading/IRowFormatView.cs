namespace OpenXmlKit.Word;

/// <summary>
/// Row formatting, as read.
/// </summary>
public interface IRowFormatView
{
    Length? Height { get; }
    RowHeightRule HeightRule { get; }

    /// <summary>
    /// Whether the row repeats at the top of every page the table spans.
    /// </summary>
    Toggle IsHeader { get; }

    Toggle CantSplit { get; }
    bool IsEmpty { get; }
}
