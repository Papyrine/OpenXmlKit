namespace OpenXmlKit.Word;

/// <summary>
/// Table formatting, as read.
/// </summary>
public interface ITableFormatView
{
    string? StyleId { get; }
    Width Width { get; }
    TableAlignment? Alignment { get; }
    TableLayout? Layout { get; }
    Length? Indent { get; }
    IBordersView Borders { get; }
    IShadingView Shading { get; }
    Length? DefaultLeftMargin { get; }
    Length? DefaultRightMargin { get; }
    Length? DefaultTopMargin { get; }
    Length? DefaultBottomMargin { get; }
    ITableLookView Look { get; }
    string? Caption { get; }
    string? Description { get; }
    bool IsEmpty { get; }
}
