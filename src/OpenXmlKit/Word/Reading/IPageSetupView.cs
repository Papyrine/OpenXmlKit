namespace OpenXmlKit.Word;

/// <summary>
/// Page geometry, as read.
/// </summary>
public interface IPageSetupView
{
    Length? PageWidth { get; }
    Length? PageHeight { get; }
    PageOrientation? Orientation { get; }
    Length? LeftMargin { get; }
    Length? RightMargin { get; }
    Length? TopMargin { get; }
    Length? BottomMargin { get; }
    Length? HeaderDistance { get; }
    Length? FooterDistance { get; }
    Length? Gutter { get; }
    SectionStart? Start { get; }
    int? ColumnCount { get; }
    Length? ColumnSpacing { get; }
    bool ColumnSeparator { get; }
    bool DifferentFirstPage { get; }
    bool DifferentOddAndEvenPages { get; }
    int? PageNumberStart { get; }
}
