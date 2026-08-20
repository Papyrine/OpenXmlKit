namespace OpenXmlKit.Word;

/// <summary>
/// Paragraph formatting: alignment, indentation, spacing, pagination, borders and shading.
/// </summary>
/// <remarks>
/// As with <see cref="Font"/>, everything defaults to unstated, so an untouched format writes no
/// paragraph properties and the style hierarchy resolves.
/// </remarks>
public partial class ParagraphFormat :
    IParagraphFormatView
{
    /// <summary>
    /// The id of the paragraph style to apply.
    /// </summary>
    /// <remarks>
    /// Worth knowing which way the precedence runs: a paragraph style beats a table style, so
    /// styling a table's font through the table style alone loses to whatever Normal says. Direct
    /// formatting on this object beats both.
    /// </remarks>
    public string? StyleId { get; set; }

    /// <summary>
    /// How the paragraph sits between its margins.
    /// </summary>
    public ParagraphAlignment? Alignment { get; set; }

    /// <summary>
    /// Space between the left margin and the paragraph.
    /// </summary>
    public Length? LeftIndent { get; set; }

    /// <summary>
    /// Space between the paragraph and the right margin.
    /// </summary>
    public Length? RightIndent { get; set; }

    /// <summary>
    /// Indents the first line further than the rest.
    /// </summary>
    public Length? FirstLineIndent { get; set; }

    /// <summary>
    /// Indents every line except the first — the shape a list marker sits in.
    /// </summary>
    public Length? HangingIndent { get; set; }

    /// <summary>
    /// Space above the paragraph.
    /// </summary>
    public Length? SpaceBefore { get; set; }

    /// <summary>
    /// Space below the paragraph.
    /// </summary>
    public Length? SpaceAfter { get; set; }

    /// <summary>
    /// Line height. Read together with <see cref="LineSpacingRule"/>: under
    /// <see cref="Word.LineSpacingRule.Multiple"/> this is a multiple of single spacing rather
    /// than a distance, so set <see cref="LineSpacingMultiple"/> instead.
    /// </summary>
    public Length? LineSpacing { get; set; }

    /// <summary>
    /// Line height as a multiple of single spacing — 1.5 for one-and-a-half spaced. Sets
    /// <see cref="LineSpacingRule"/> to <see cref="Word.LineSpacingRule.Multiple"/>.
    /// </summary>
    public double? LineSpacingMultiple { get; set; }

    /// <summary>
    /// How <see cref="LineSpacing"/> is interpreted.
    /// </summary>
    public LineSpacingRule? LineSpacingRule { get; set; }

    /// <summary>
    /// Suppresses <see cref="SpaceBefore"/> and <see cref="SpaceAfter"/> between consecutive
    /// paragraphs of the same style — what keeps a bulleted list from being double spaced.
    /// </summary>
    public Toggle ContextualSpacing { get; set; }

    /// <summary>
    /// Keeps this paragraph on the same page as the one after it.
    /// </summary>
    public Toggle KeepWithNext { get; set; }

    /// <summary>
    /// Keeps this paragraph from being split across a page boundary.
    /// </summary>
    public Toggle KeepTogether { get; set; }

    /// <summary>
    /// Starts the paragraph on a new page.
    /// </summary>
    public Toggle PageBreakBefore { get; set; }

    /// <summary>
    /// Prevents a single line being stranded at the top or bottom of a page.
    /// </summary>
    public Toggle WidowControl { get; set; }

    /// <summary>
    /// Leaves the paragraph out of the line numbering, where a section has it on.
    /// </summary>
    public Toggle SuppressLineNumbers { get; set; }

    /// <summary>
    /// Lays the paragraph out right to left.
    /// </summary>
    public Toggle RightToLeft { get; set; }

    /// <summary>
    /// The heading level this paragraph contributes to the navigation pane and a table of
    /// contents. 0 is the top level; 9 means body text.
    /// </summary>
    public int? OutlineLevel { get; set; }

    /// <summary>
    /// Rules drawn around the paragraph.
    /// </summary>
    public Borders Borders { get; } = new();

    /// <summary>
    /// A fill behind the whole paragraph.
    /// </summary>
    public Shading Shading { get; } = new();

    /// <summary>
    /// Where tabs in the paragraph land.
    /// </summary>
    public TabStops TabStops { get; } = new();

    /// <summary>
    /// The list this paragraph belongs to, if any.
    /// </summary>
    public ListMembership? List { get; set; }

    /// <summary>
    /// Whether anything at all is stated. An empty format writes no properties element,
    /// leaving the style hierarchy to resolve every value.
    /// </summary>
    public bool IsEmpty =>
        StyleId == null &&
        Alignment == null &&
        LeftIndent == null &&
        RightIndent == null &&
        FirstLineIndent == null &&
        HangingIndent == null &&
        SpaceBefore == null &&
        SpaceAfter == null &&
        LineSpacing == null &&
        LineSpacingMultiple == null &&
        LineSpacingRule == null &&
        !ContextualSpacing.IsSet &&
        !KeepWithNext.IsSet &&
        !KeepTogether.IsSet &&
        !PageBreakBefore.IsSet &&
        !WidowControl.IsSet &&
        !SuppressLineNumbers.IsSet &&
        !RightToLeft.IsSet &&
        OutlineLevel == null &&
        Borders.IsEmpty &&
        Shading.IsEmpty &&
        TabStops.IsEmpty &&
        List == null;

    /// <summary>
    /// An independent copy, so the two can diverge.
    /// </summary>
    public ParagraphFormat Clone()
    {
        var clone = new ParagraphFormat();
        clone.CopyFrom(this);
        return clone;
    }

    /// <summary>
    /// Overwrites every property with the other value, stated or not.
    /// </summary>
    public void CopyFrom(ParagraphFormat other)
    {
        StyleId = other.StyleId;
        Alignment = other.Alignment;
        LeftIndent = other.LeftIndent;
        RightIndent = other.RightIndent;
        FirstLineIndent = other.FirstLineIndent;
        HangingIndent = other.HangingIndent;
        SpaceBefore = other.SpaceBefore;
        SpaceAfter = other.SpaceAfter;
        LineSpacing = other.LineSpacing;
        LineSpacingMultiple = other.LineSpacingMultiple;
        LineSpacingRule = other.LineSpacingRule;
        ContextualSpacing = other.ContextualSpacing;
        KeepWithNext = other.KeepWithNext;
        KeepTogether = other.KeepTogether;
        PageBreakBefore = other.PageBreakBefore;
        WidowControl = other.WidowControl;
        SuppressLineNumbers = other.SuppressLineNumbers;
        RightToLeft = other.RightToLeft;
        OutlineLevel = other.OutlineLevel;
        Borders.CopyFrom(other.Borders);
        Shading.CopyFrom(other.Shading);
        TabStops.CopyFrom(other.TabStops);
        List = other.List;
    }

    /// <summary>
    /// Returns every property to unstated, so the style hierarchy resolves the lot.
    /// </summary>
    public void Clear() =>
        CopyFrom(new());

    internal W.ParagraphProperties? ToProperties(Font? runFont = null)
    {
        if (IsEmpty &&
            runFont is not { IsEmpty: false })
        {
            return null;
        }

        var properties = new W.ParagraphProperties();
        ApplyTo(properties, runFont);
        return properties;
    }

    /// <summary>
    /// Fills the paragraph properties. Children go in through typed properties, which is what
    /// places each at its position in the CT_PPr sequence.
    /// </summary>
    /// <param name="properties">
    /// The element to fill.
    /// </param>
    /// <param name="runFont">
    /// Formatting for the paragraph mark itself — the pilcrow — which lives inside pPr rather than
    /// beside it. It is what decides the height of an empty paragraph, and what a following
    /// paragraph inherits when this one is split.
    /// </param>
    internal void ApplyTo(W.ParagraphProperties properties, Font? runFont = null)
    {
        if (StyleId != null)
        {
            properties.ParagraphStyleId = new()
            {
                Val = StyleId
            };
        }

        if (List is { } list)
        {
            properties.NumberingProperties = new()
            {
                NumberingLevelReference = new()
                {
                    Val = list.Level
                },
                NumberingId = new()
                {
                    Val = list.NumberingId
                }
            };
        }

        properties.SuppressLineNumbers = Toggles.OnOff<W.SuppressLineNumbers>(SuppressLineNumbers);

        if (Borders.ToParagraphBorders() is { } borders)
        {
            properties.ParagraphBorders = borders;
        }

        if (Shading.ToOpenXml() is { } shading)
        {
            properties.Shading = shading;
        }

        if (TabStops.ToOpenXml() is { } tabs)
        {
            properties.Tabs = tabs;
        }

        properties.KeepNext = Toggles.OnOff<W.KeepNext>(KeepWithNext);
        properties.KeepLines = Toggles.OnOff<W.KeepLines>(KeepTogether);
        properties.PageBreakBefore = Toggles.OnOff<W.PageBreakBefore>(PageBreakBefore);
        properties.WidowControl = Toggles.OnOff<W.WidowControl>(WidowControl);
        properties.ContextualSpacing = Toggles.OnOff<W.ContextualSpacing>(ContextualSpacing);
        properties.BiDi = Toggles.OnOff<W.BiDi>(RightToLeft);

        ApplySpacing(properties);
        ApplyIndentation(properties);

        if (Alignment is { } alignment)
        {
            properties.Justification = new()
            {
                Val = alignment.ToOpenXml()
            };
        }

        if (OutlineLevel is { } outlineLevel)
        {
            properties.OutlineLevel = new()
            {
                Val = outlineLevel
            };
        }

        if (runFont is { IsEmpty: false })
        {
            var markProperties = new W.ParagraphMarkRunProperties();
            // The paragraph mark carries its own rPr, whose children follow the same CT_RPr
            // sequence. Building a RunProperties first and moving the children across reuses the
            // ordering the SDK applies there rather than restating it.
            var source = runFont.ToProperties();
            if (source != null)
            {
                foreach (var child in source.ChildElements.ToList())
                {
                    child.Remove();
                    markProperties.AppendChild(child);
                }
            }

            properties.ParagraphMarkRunProperties = markProperties;
        }
    }

    void ApplySpacing(W.ParagraphProperties properties)
    {
        if (SpaceBefore == null &&
            SpaceAfter == null &&
            LineSpacing == null &&
            LineSpacingMultiple == null &&
            LineSpacingRule == null)
        {
            return;
        }

        var spacing = new W.SpacingBetweenLines();
        if (SpaceBefore is { } before)
        {
            spacing.Before = before.Twips.ToString(CultureInfo.InvariantCulture);
        }

        if (SpaceAfter is { } after)
        {
            spacing.After = after.Twips.ToString(CultureInfo.InvariantCulture);
        }

        if (LineSpacingMultiple is { } multiple)
        {
            // Under the "auto" rule, w:line is measured in 240ths of a line rather than in twips,
            // so single spacing is 240 and one-and-a-half is 360.
            spacing.Line = ((int) Math.Round(multiple * 240)).ToString(CultureInfo.InvariantCulture);
            spacing.LineRule = W.LineSpacingRuleValues.Auto;
        }
        else if (LineSpacing is { } line)
        {
            spacing.Line = line.Twips.ToString(CultureInfo.InvariantCulture);
            spacing.LineRule = (LineSpacingRule ?? Word.LineSpacingRule.AtLeast).ToOpenXml();
        }
        else if (LineSpacingRule is { } rule)
        {
            spacing.LineRule = rule.ToOpenXml();
        }

        properties.SpacingBetweenLines = spacing;
    }

    void ApplyIndentation(W.ParagraphProperties properties)
    {
        if (LeftIndent == null &&
            RightIndent == null &&
            FirstLineIndent == null &&
            HangingIndent == null)
        {
            return;
        }

        var indentation = new W.Indentation();
        if (LeftIndent is { } left)
        {
            indentation.Left = left.Twips.ToString(CultureInfo.InvariantCulture);
        }

        if (RightIndent is { } right)
        {
            indentation.Right = right.Twips.ToString(CultureInfo.InvariantCulture);
        }

        // Hanging and first-line are the same attribute slot pulling in opposite directions, and
        // Word honours hanging when both are present. Stating only one avoids the ambiguity.
        if (HangingIndent is { } hanging)
        {
            indentation.Hanging = hanging.Twips.ToString(CultureInfo.InvariantCulture);
        }
        else if (FirstLineIndent is { } firstLine)
        {
            indentation.FirstLine = firstLine.Twips.ToString(CultureInfo.InvariantCulture);
        }

        properties.Indentation = indentation;
    }


    internal void ReadFrom(W.ParagraphProperties properties)
    {
        StyleId = properties.ParagraphStyleId?.Val?.Value;

        if (properties.NumberingProperties is { NumberingId.Val: { HasValue: true } numberingId } numbering)
        {
            List = new(numberingId.Value, numbering.NumberingLevelReference?.Val?.Value ?? 0);
        }

        if (properties.Justification?.Val is { HasValue: true } justification)
        {
            Alignment = Map.ToAlignment(justification.Value);
        }

        KeepWithNext = Toggles.Read(properties.KeepNext);
        KeepTogether = Toggles.Read(properties.KeepLines);
        PageBreakBefore = Toggles.Read(properties.PageBreakBefore);
        WidowControl = Toggles.Read(properties.WidowControl);
        ContextualSpacing = Toggles.Read(properties.ContextualSpacing);
        SuppressLineNumbers = Toggles.Read(properties.SuppressLineNumbers);
        RightToLeft = Toggles.Read(properties.BiDi);

        if (properties.OutlineLevel?.Val is { HasValue: true } outlineLevel)
        {
            OutlineLevel = outlineLevel.Value;
        }

        if (properties.SpacingBetweenLines is { } spacing)
        {
            SpaceBefore = ReadTwips(spacing.Before?.Value);
            SpaceAfter = ReadTwips(spacing.After?.Value);

            var rule = spacing.LineRule?.Value;
            if (rule != null)
            {
                LineSpacingRule = Map.ToLineSpacingRule(rule.Value);
            }

            if (spacing.Line?.Value is { } line &&
                double.TryParse(line, NumberStyles.Float, CultureInfo.InvariantCulture, out var lineValue))
            {
                if (rule == null ||
                    rule.Value == W.LineSpacingRuleValues.Auto)
                {
                    LineSpacingMultiple = lineValue / 240;
                }
                else
                {
                    LineSpacing = Length.FromTwips(lineValue);
                }
            }
        }

        if (properties.Indentation is { } indentation)
        {
            LeftIndent = ReadTwips(indentation.Left?.Value);
            RightIndent = ReadTwips(indentation.Right?.Value);
            FirstLineIndent = ReadTwips(indentation.FirstLine?.Value);
            HangingIndent = ReadTwips(indentation.Hanging?.Value);
        }

        Borders.ReadFrom(properties.ParagraphBorders);
        Shading.ReadFrom(properties.Shading);
        TabStops.ReadFrom(properties.Tabs);
    }

    static Length? ReadTwips(string? value)
    {
        if (value == null ||
            !double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var twips))
        {
            return null;
        }

        return Length.FromTwips(twips);
    }

}
