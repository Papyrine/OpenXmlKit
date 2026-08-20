namespace OpenXmlKit.Word;

/// <summary>
/// Paragraph formatting, as read.
/// </summary>
public interface IParagraphFormatView
{
    string? StyleId { get; }
    ParagraphAlignment? Alignment { get; }
    Length? LeftIndent { get; }
    Length? RightIndent { get; }
    Length? FirstLineIndent { get; }
    Length? HangingIndent { get; }
    Length? SpaceBefore { get; }
    Length? SpaceAfter { get; }
    Length? LineSpacing { get; }
    double? LineSpacingMultiple { get; }
    LineSpacingRule? LineSpacingRule { get; }
    Toggle ContextualSpacing { get; }
    Toggle KeepWithNext { get; }
    Toggle KeepTogether { get; }
    Toggle PageBreakBefore { get; }
    Toggle WidowControl { get; }
    Toggle SuppressLineNumbers { get; }
    Toggle RightToLeft { get; }
    int? OutlineLevel { get; }
    IBordersView Borders { get; }
    IShadingView Shading { get; }
    IReadOnlyList<TabStop> TabStops { get; }

    /// <summary>
    /// The list this paragraph belongs to, if any.
    /// </summary>
    ListMembership? List { get; }

    bool IsEmpty { get; }
}
