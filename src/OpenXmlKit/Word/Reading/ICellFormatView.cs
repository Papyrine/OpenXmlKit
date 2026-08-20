namespace OpenXmlKit.Word;

/// <summary>
/// Cell formatting, as read.
/// </summary>
public interface ICellFormatView
{
    /// <inheritdoc cref="CellFormat.Width"/>
    Width Width { get; }

    /// <inheritdoc cref="CellFormat.ColumnSpan"/>
    int ColumnSpan { get; }

    /// <inheritdoc cref="CellFormat.VerticalMerge"/>
    CellMerge VerticalMerge { get; }

    /// <inheritdoc cref="CellFormat.Borders"/>
    IBordersView Borders { get; }

    /// <inheritdoc cref="CellFormat.Shading"/>
    IShadingView Shading { get; }

    /// <inheritdoc cref="CellFormat.VerticalAlignment"/>
    VerticalAlignment? VerticalAlignment { get; }

    /// <inheritdoc cref="CellFormat.TextDirection"/>
    TextDirection? TextDirection { get; }

    /// <inheritdoc cref="CellFormat.LeftMargin"/>
    Length? LeftMargin { get; }

    /// <inheritdoc cref="CellFormat.RightMargin"/>
    Length? RightMargin { get; }

    /// <inheritdoc cref="CellFormat.TopMargin"/>
    Length? TopMargin { get; }

    /// <inheritdoc cref="CellFormat.BottomMargin"/>
    Length? BottomMargin { get; }

    /// <inheritdoc cref="CellFormat.NoWrap"/>
    Toggle NoWrap { get; }

    /// <inheritdoc cref="CellFormat.FitText"/>
    Toggle FitText { get; }

    /// <inheritdoc cref="CellFormat.IsEmpty"/>
    bool IsEmpty { get; }
}
