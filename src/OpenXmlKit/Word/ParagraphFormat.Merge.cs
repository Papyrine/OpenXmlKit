namespace OpenXmlKit.Word;

public partial class ParagraphFormat
{
    /// <summary>
    /// Overlays another paragraph format on this one: everything it states wins, everything it
    /// leaves unstated is left alone.
    /// </summary>
    public void MergeFrom(ParagraphFormat higher)
    {
        Alignment = higher.Alignment ?? Alignment;
        LeftIndent = higher.LeftIndent ?? LeftIndent;
        RightIndent = higher.RightIndent ?? RightIndent;
        FirstLineIndent = higher.FirstLineIndent ?? FirstLineIndent;
        HangingIndent = higher.HangingIndent ?? HangingIndent;
        SpaceBefore = higher.SpaceBefore ?? SpaceBefore;
        SpaceAfter = higher.SpaceAfter ?? SpaceAfter;
        LineSpacing = higher.LineSpacing ?? LineSpacing;
        LineSpacingMultiple = higher.LineSpacingMultiple ?? LineSpacingMultiple;
        LineSpacingRule = higher.LineSpacingRule ?? LineSpacingRule;
        OutlineLevel = higher.OutlineLevel ?? OutlineLevel;
        List = higher.List ?? List;

        ContextualSpacing = Overlay(ContextualSpacing, higher.ContextualSpacing);
        KeepWithNext = Overlay(KeepWithNext, higher.KeepWithNext);
        KeepTogether = Overlay(KeepTogether, higher.KeepTogether);
        PageBreakBefore = Overlay(PageBreakBefore, higher.PageBreakBefore);
        WidowControl = Overlay(WidowControl, higher.WidowControl);
        SuppressLineNumbers = Overlay(SuppressLineNumbers, higher.SuppressLineNumbers);
        RightToLeft = Overlay(RightToLeft, higher.RightToLeft);

        if (!higher.Borders.IsEmpty)
        {
            Borders.CopyFrom(higher.Borders);
        }

        if (!higher.Shading.IsEmpty)
        {
            Shading.CopyFrom(higher.Shading);
        }

        if (!higher.TabStops.IsEmpty)
        {
            foreach (var stop in higher.TabStops)
            {
                TabStops.Add(stop);
            }
        }
    }

    static Toggle Overlay(Toggle lower, Toggle higher) =>
        higher.IsSet ? higher : lower;
}
