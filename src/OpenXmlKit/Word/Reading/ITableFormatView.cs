namespace OpenXmlKit.Word;

/// <summary>
/// Table formatting, as read.
/// </summary>
public interface ITableFormatView
{
    /// <inheritdoc cref="TableFormat.StyleId"/>
    string? StyleId { get; }

    /// <inheritdoc cref="TableFormat.Width"/>
    Width Width { get; }

    /// <inheritdoc cref="TableFormat.Alignment"/>
    TableAlignment? Alignment { get; }

    /// <inheritdoc cref="TableFormat.Layout"/>
    TableLayout? Layout { get; }

    /// <inheritdoc cref="TableFormat.Indent"/>
    Length? Indent { get; }

    /// <inheritdoc cref="TableFormat.Borders"/>
    IBordersView Borders { get; }

    /// <inheritdoc cref="TableFormat.Shading"/>
    IShadingView Shading { get; }

    /// <inheritdoc cref="TableFormat.DefaultLeftMargin"/>
    Length? DefaultLeftMargin { get; }

    /// <inheritdoc cref="TableFormat.DefaultRightMargin"/>
    Length? DefaultRightMargin { get; }

    /// <inheritdoc cref="TableFormat.DefaultTopMargin"/>
    Length? DefaultTopMargin { get; }

    /// <inheritdoc cref="TableFormat.DefaultBottomMargin"/>
    Length? DefaultBottomMargin { get; }

    /// <inheritdoc cref="TableFormat.Look"/>
    ITableLookView Look { get; }

    /// <inheritdoc cref="TableFormat.Caption"/>
    string? Caption { get; }

    /// <inheritdoc cref="TableFormat.Description"/>
    string? Description { get; }

    /// <inheritdoc cref="TableFormat.IsEmpty"/>
    bool IsEmpty { get; }
}
