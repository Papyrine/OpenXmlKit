namespace OpenXmlKit.Word;

/// <summary>
/// Cell formatting, as read.
/// </summary>
public interface ICellFormatView
{
    Width Width { get; }
    int ColumnSpan { get; }
    CellMerge VerticalMerge { get; }
    IBordersView Borders { get; }
    IShadingView Shading { get; }
    VerticalAlignment? VerticalAlignment { get; }
    TextDirection? TextDirection { get; }
    Length? LeftMargin { get; }
    Length? RightMargin { get; }
    Length? TopMargin { get; }
    Length? BottomMargin { get; }
    Toggle NoWrap { get; }
    Toggle FitText { get; }
    bool IsEmpty { get; }
}
