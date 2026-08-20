namespace OpenXmlKit.Word;

/// <summary>
/// Row formatting, as read.
/// </summary>
public interface IRowFormatView
{
    /// <inheritdoc cref="RowFormat.Height"/>
    Length? Height { get; }

    /// <inheritdoc cref="RowFormat.HeightRule"/>
    RowHeightRule HeightRule { get; }

    /// <summary>
    /// Whether the row repeats at the top of every page the table spans.
    /// </summary>
    Toggle IsHeader { get; }

    /// <inheritdoc cref="RowFormat.CantSplit"/>
    Toggle CantSplit { get; }

    /// <inheritdoc cref="RowFormat.GridBefore"/>
    int? GridBefore { get; }

    /// <inheritdoc cref="RowFormat.GridAfter"/>
    int? GridAfter { get; }

    /// <inheritdoc cref="RowFormat.WidthBefore"/>
    Width WidthBefore { get; }

    /// <inheritdoc cref="RowFormat.WidthAfter"/>
    Width WidthAfter { get; }

    /// <inheritdoc cref="RowFormat.IsEmpty"/>
    bool IsEmpty { get; }
}
