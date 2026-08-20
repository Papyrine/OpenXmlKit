namespace OpenXmlKit.Word;

/// <summary>
/// Paragraph formatting, as read.
/// </summary>
public interface IParagraphFormatView
{
    /// <inheritdoc cref="ParagraphFormat.StyleId"/>
    string? StyleId { get; }

    /// <inheritdoc cref="ParagraphFormat.Alignment"/>
    ParagraphAlignment? Alignment { get; }

    /// <inheritdoc cref="ParagraphFormat.LeftIndent"/>
    Length? LeftIndent { get; }

    /// <inheritdoc cref="ParagraphFormat.RightIndent"/>
    Length? RightIndent { get; }

    /// <inheritdoc cref="ParagraphFormat.FirstLineIndent"/>
    Length? FirstLineIndent { get; }

    /// <inheritdoc cref="ParagraphFormat.HangingIndent"/>
    Length? HangingIndent { get; }

    /// <inheritdoc cref="ParagraphFormat.SpaceBefore"/>
    Length? SpaceBefore { get; }

    /// <inheritdoc cref="ParagraphFormat.SpaceAfter"/>
    Length? SpaceAfter { get; }

    /// <inheritdoc cref="ParagraphFormat.LineSpacing"/>
    Length? LineSpacing { get; }

    /// <inheritdoc cref="ParagraphFormat.LineSpacingMultiple"/>
    double? LineSpacingMultiple { get; }

    /// <inheritdoc cref="ParagraphFormat.LineSpacingRule"/>
    LineSpacingRule? LineSpacingRule { get; }

    /// <inheritdoc cref="ParagraphFormat.ContextualSpacing"/>
    Toggle ContextualSpacing { get; }

    /// <inheritdoc cref="ParagraphFormat.KeepWithNext"/>
    Toggle KeepWithNext { get; }

    /// <inheritdoc cref="ParagraphFormat.KeepTogether"/>
    Toggle KeepTogether { get; }

    /// <inheritdoc cref="ParagraphFormat.PageBreakBefore"/>
    Toggle PageBreakBefore { get; }

    /// <inheritdoc cref="ParagraphFormat.WidowControl"/>
    Toggle WidowControl { get; }

    /// <inheritdoc cref="ParagraphFormat.SuppressLineNumbers"/>
    Toggle SuppressLineNumbers { get; }

    /// <inheritdoc cref="ParagraphFormat.RightToLeft"/>
    Toggle RightToLeft { get; }

    /// <inheritdoc cref="ParagraphFormat.OutlineLevel"/>
    int? OutlineLevel { get; }

    /// <inheritdoc cref="ParagraphFormat.Borders"/>
    IBordersView Borders { get; }

    /// <inheritdoc cref="ParagraphFormat.Shading"/>
    IShadingView Shading { get; }

    /// <inheritdoc cref="ParagraphFormat.TabStops"/>
    IReadOnlyList<TabStop> TabStops { get; }

    /// <summary>
    /// The list this paragraph belongs to, if any.
    /// </summary>
    ListMembership? List { get; }

    /// <inheritdoc cref="ParagraphFormat.IsEmpty"/>
    bool IsEmpty { get; }
}
