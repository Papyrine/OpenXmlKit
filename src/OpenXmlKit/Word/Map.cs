namespace OpenXmlKit.Word;

/// <summary>
/// Translates this library's enums to the SDK's, and back.
/// </summary>
/// <remarks>
/// The SDK models its enumerations as structs implementing IEnumValue rather than as CLR enums,
/// which means they cannot appear in a switch arm, carry no attributes, and do not round-trip
/// through Enum.Parse. Declaring real enums for the public API and mapping here keeps all of that
/// available to callers; the cost is this file.
/// </remarks>
static class Map
{
    public static W.JustificationValues ToOpenXml(this ParagraphAlignment value) =>
        value switch
        {
            ParagraphAlignment.Left => W.JustificationValues.Left,
            ParagraphAlignment.Center => W.JustificationValues.Center,
            ParagraphAlignment.Right => W.JustificationValues.Right,
            // Word calls justified text "both", as in flush to both margins.
            ParagraphAlignment.Justify => W.JustificationValues.Both,
            ParagraphAlignment.Distribute => W.JustificationValues.Distribute,
            _ => throw Unmapped(value)
        };

    public static ParagraphAlignment ToAlignment(W.JustificationValues value)
    {
        if (value == W.JustificationValues.Center)
        {
            return ParagraphAlignment.Center;
        }

        if (value == W.JustificationValues.Right)
        {
            return ParagraphAlignment.Right;
        }

        if (value == W.JustificationValues.Both)
        {
            return ParagraphAlignment.Justify;
        }

        if (value == W.JustificationValues.Distribute)
        {
            return ParagraphAlignment.Distribute;
        }

        return ParagraphAlignment.Left;
    }

    public static W.UnderlineValues ToOpenXml(this UnderlineStyle value) =>
        value switch
        {
            UnderlineStyle.None => W.UnderlineValues.None,
            UnderlineStyle.Single => W.UnderlineValues.Single,
            UnderlineStyle.Double => W.UnderlineValues.Double,
            UnderlineStyle.Thick => W.UnderlineValues.Thick,
            UnderlineStyle.Dotted => W.UnderlineValues.Dotted,
            UnderlineStyle.Dashed => W.UnderlineValues.Dash,
            UnderlineStyle.DotDash => W.UnderlineValues.DotDash,
            UnderlineStyle.DotDotDash => W.UnderlineValues.DotDotDash,
            UnderlineStyle.Wave => W.UnderlineValues.Wave,
            UnderlineStyle.Words => W.UnderlineValues.Words,
            _ => throw Unmapped(value)
        };

    public static UnderlineStyle ToUnderline(W.UnderlineValues value)
    {
        if (value == W.UnderlineValues.Single)
        {
            return UnderlineStyle.Single;
        }

        if (value == W.UnderlineValues.Double)
        {
            return UnderlineStyle.Double;
        }

        if (value == W.UnderlineValues.Thick)
        {
            return UnderlineStyle.Thick;
        }

        if (value == W.UnderlineValues.Dotted)
        {
            return UnderlineStyle.Dotted;
        }

        if (value == W.UnderlineValues.Dash)
        {
            return UnderlineStyle.Dashed;
        }

        if (value == W.UnderlineValues.DotDash)
        {
            return UnderlineStyle.DotDash;
        }

        if (value == W.UnderlineValues.DotDotDash)
        {
            return UnderlineStyle.DotDotDash;
        }

        if (value == W.UnderlineValues.Wave)
        {
            return UnderlineStyle.Wave;
        }

        if (value == W.UnderlineValues.Words)
        {
            return UnderlineStyle.Words;
        }

        return UnderlineStyle.None;
    }

    public static W.VerticalPositionValues ToOpenXml(this VerticalTextPosition value) =>
        value switch
        {
            VerticalTextPosition.Superscript => W.VerticalPositionValues.Superscript,
            VerticalTextPosition.Subscript => W.VerticalPositionValues.Subscript,
            _ => W.VerticalPositionValues.Baseline
        };

    public static VerticalTextPosition ToVerticalTextPosition(W.VerticalPositionValues value)
    {
        if (value == W.VerticalPositionValues.Superscript)
        {
            return VerticalTextPosition.Superscript;
        }

        if (value == W.VerticalPositionValues.Subscript)
        {
            return VerticalTextPosition.Subscript;
        }

        return VerticalTextPosition.Baseline;
    }

    public static W.BorderValues ToOpenXml(this BorderStyle value) =>
        value switch
        {
            // "nil" rather than "none": both suppress the border, but only nil overrides one
            // inherited from a table style. none reads as "nothing stated" to some consumers.
            BorderStyle.None => W.BorderValues.Nil,
            BorderStyle.Single => W.BorderValues.Single,
            BorderStyle.Thick => W.BorderValues.Thick,
            BorderStyle.Double => W.BorderValues.Double,
            BorderStyle.Dotted => W.BorderValues.Dotted,
            BorderStyle.Dashed => W.BorderValues.Dashed,
            BorderStyle.DotDash => W.BorderValues.DotDash,
            BorderStyle.DotDotDash => W.BorderValues.DotDotDash,
            BorderStyle.Triple => W.BorderValues.Triple,
            BorderStyle.Wave => W.BorderValues.Wave,
            BorderStyle.DoubleWave => W.BorderValues.DoubleWave,
            BorderStyle.Inset => W.BorderValues.Inset,
            BorderStyle.Outset => W.BorderValues.Outset,
            _ => throw Unmapped(value)
        };

    public static BorderStyle ToBorderStyle(W.BorderValues value)
    {
        if (value == W.BorderValues.Single)
        {
            return BorderStyle.Single;
        }

        if (value == W.BorderValues.Thick)
        {
            return BorderStyle.Thick;
        }

        if (value == W.BorderValues.Double)
        {
            return BorderStyle.Double;
        }

        if (value == W.BorderValues.Dotted)
        {
            return BorderStyle.Dotted;
        }

        if (value == W.BorderValues.Dashed)
        {
            return BorderStyle.Dashed;
        }

        if (value == W.BorderValues.DotDash)
        {
            return BorderStyle.DotDash;
        }

        if (value == W.BorderValues.DotDotDash)
        {
            return BorderStyle.DotDotDash;
        }

        if (value == W.BorderValues.Triple)
        {
            return BorderStyle.Triple;
        }

        if (value == W.BorderValues.Wave)
        {
            return BorderStyle.Wave;
        }

        if (value == W.BorderValues.DoubleWave)
        {
            return BorderStyle.DoubleWave;
        }

        if (value == W.BorderValues.Inset)
        {
            return BorderStyle.Inset;
        }

        if (value == W.BorderValues.Outset)
        {
            return BorderStyle.Outset;
        }

        return BorderStyle.None;
    }

    public static W.TableVerticalAlignmentValues ToOpenXml(this VerticalAlignment value) =>
        value switch
        {
            VerticalAlignment.Center => W.TableVerticalAlignmentValues.Center,
            VerticalAlignment.Bottom => W.TableVerticalAlignmentValues.Bottom,
            _ => W.TableVerticalAlignmentValues.Top
        };

    public static VerticalAlignment ToVerticalAlignment(W.TableVerticalAlignmentValues value)
    {
        if (value == W.TableVerticalAlignmentValues.Center)
        {
            return VerticalAlignment.Center;
        }

        if (value == W.TableVerticalAlignmentValues.Bottom)
        {
            return VerticalAlignment.Bottom;
        }

        return VerticalAlignment.Top;
    }

    public static W.LineSpacingRuleValues ToOpenXml(this LineSpacingRule value) =>
        value switch
        {
            LineSpacingRule.AtLeast => W.LineSpacingRuleValues.AtLeast,
            LineSpacingRule.Exactly => W.LineSpacingRuleValues.Exact,
            _ => W.LineSpacingRuleValues.Auto
        };

    public static LineSpacingRule ToLineSpacingRule(W.LineSpacingRuleValues value)
    {
        if (value == W.LineSpacingRuleValues.AtLeast)
        {
            return LineSpacingRule.AtLeast;
        }

        if (value == W.LineSpacingRuleValues.Exact)
        {
            return LineSpacingRule.Exactly;
        }

        return LineSpacingRule.Multiple;
    }

    public static W.HeightRuleValues ToOpenXml(this RowHeightRule value) =>
        value switch
        {
            RowHeightRule.AtLeast => W.HeightRuleValues.AtLeast,
            RowHeightRule.Exactly => W.HeightRuleValues.Exact,
            _ => W.HeightRuleValues.Auto
        };

    public static RowHeightRule ToRowHeightRule(W.HeightRuleValues value)
    {
        if (value == W.HeightRuleValues.AtLeast)
        {
            return RowHeightRule.AtLeast;
        }

        if (value == W.HeightRuleValues.Exact)
        {
            return RowHeightRule.Exactly;
        }

        return RowHeightRule.Auto;
    }

    public static W.StyleValues ToOpenXml(this StyleKind value) =>
        value switch
        {
            StyleKind.Character => W.StyleValues.Character,
            StyleKind.Table => W.StyleValues.Table,
            StyleKind.Numbering => W.StyleValues.Numbering,
            _ => W.StyleValues.Paragraph
        };

    public static StyleKind ToStyleKind(W.StyleValues value)
    {
        if (value == W.StyleValues.Character)
        {
            return StyleKind.Character;
        }

        if (value == W.StyleValues.Table)
        {
            return StyleKind.Table;
        }

        if (value == W.StyleValues.Numbering)
        {
            return StyleKind.Numbering;
        }

        return StyleKind.Paragraph;
    }

    public static W.HeaderFooterValues ToOpenXml(this HeaderFooterKind value) =>
        value switch
        {
            HeaderFooterKind.First => W.HeaderFooterValues.First,
            HeaderFooterKind.Even => W.HeaderFooterValues.Even,
            _ => W.HeaderFooterValues.Default
        };

    public static HeaderFooterKind ToHeaderFooterKind(W.HeaderFooterValues value)
    {
        if (value == W.HeaderFooterValues.First)
        {
            return HeaderFooterKind.First;
        }

        if (value == W.HeaderFooterValues.Even)
        {
            return HeaderFooterKind.Even;
        }

        return HeaderFooterKind.Default;
    }

    public static W.PageOrientationValues ToOpenXml(this PageOrientation value) =>
        value == PageOrientation.Landscape
            ? W.PageOrientationValues.Landscape
            : W.PageOrientationValues.Portrait;

    public static PageOrientation ToOrientation(W.PageOrientationValues value) =>
        value == W.PageOrientationValues.Landscape
            ? PageOrientation.Landscape
            : PageOrientation.Portrait;

    public static W.SectionMarkValues ToOpenXml(this SectionStart value) =>
        value switch
        {
            SectionStart.Continuous => W.SectionMarkValues.Continuous,
            SectionStart.NewColumn => W.SectionMarkValues.NextColumn,
            SectionStart.EvenPage => W.SectionMarkValues.EvenPage,
            SectionStart.OddPage => W.SectionMarkValues.OddPage,
            _ => W.SectionMarkValues.NextPage
        };

    public static SectionStart ToSectionStart(W.SectionMarkValues value)
    {
        if (value == W.SectionMarkValues.Continuous)
        {
            return SectionStart.Continuous;
        }

        if (value == W.SectionMarkValues.NextColumn)
        {
            return SectionStart.NewColumn;
        }

        if (value == W.SectionMarkValues.EvenPage)
        {
            return SectionStart.EvenPage;
        }

        if (value == W.SectionMarkValues.OddPage)
        {
            return SectionStart.OddPage;
        }

        return SectionStart.NewPage;
    }

    public static W.TableLayoutValues ToOpenXml(this TableLayout value) =>
        value == TableLayout.Fixed
            ? W.TableLayoutValues.Fixed
            : W.TableLayoutValues.Autofit;

    public static W.MergedCellValues ToOpenXml(this CellMerge value) =>
        value == CellMerge.Restart
            ? W.MergedCellValues.Restart
            : W.MergedCellValues.Continue;

    public static W.ThemeColorValues ToOpenXml(this ThemeColor value) =>
        value switch
        {
            ThemeColor.Background1 => W.ThemeColorValues.Background1,
            ThemeColor.Text1 => W.ThemeColorValues.Text1,
            ThemeColor.Background2 => W.ThemeColorValues.Background2,
            ThemeColor.Text2 => W.ThemeColorValues.Text2,
            ThemeColor.Accent1 => W.ThemeColorValues.Accent1,
            ThemeColor.Accent2 => W.ThemeColorValues.Accent2,
            ThemeColor.Accent3 => W.ThemeColorValues.Accent3,
            ThemeColor.Accent4 => W.ThemeColorValues.Accent4,
            ThemeColor.Accent5 => W.ThemeColorValues.Accent5,
            ThemeColor.Accent6 => W.ThemeColorValues.Accent6,
            ThemeColor.Hyperlink => W.ThemeColorValues.Hyperlink,
            ThemeColor.FollowedHyperlink => W.ThemeColorValues.FollowedHyperlink,
            _ => throw Unmapped(value)
        };

    public static W.HighlightColorValues ToOpenXml(this HighlightColor value) =>
        value switch
        {
            HighlightColor.Black => W.HighlightColorValues.Black,
            HighlightColor.Blue => W.HighlightColorValues.Blue,
            HighlightColor.Cyan => W.HighlightColorValues.Cyan,
            HighlightColor.Green => W.HighlightColorValues.Green,
            HighlightColor.Magenta => W.HighlightColorValues.Magenta,
            HighlightColor.Red => W.HighlightColorValues.Red,
            HighlightColor.Yellow => W.HighlightColorValues.Yellow,
            HighlightColor.White => W.HighlightColorValues.White,
            HighlightColor.DarkBlue => W.HighlightColorValues.DarkBlue,
            HighlightColor.DarkCyan => W.HighlightColorValues.DarkCyan,
            HighlightColor.DarkGreen => W.HighlightColorValues.DarkGreen,
            HighlightColor.DarkMagenta => W.HighlightColorValues.DarkMagenta,
            HighlightColor.DarkRed => W.HighlightColorValues.DarkRed,
            HighlightColor.DarkYellow => W.HighlightColorValues.DarkYellow,
            HighlightColor.DarkGray => W.HighlightColorValues.DarkGray,
            HighlightColor.LightGray => W.HighlightColorValues.LightGray,
            _ => W.HighlightColorValues.None
        };

    public static HighlightColor ToHighlight(W.HighlightColorValues value)
    {
        if (value == W.HighlightColorValues.Black)
        {
            return HighlightColor.Black;
        }

        if (value == W.HighlightColorValues.Blue)
        {
            return HighlightColor.Blue;
        }

        if (value == W.HighlightColorValues.Cyan)
        {
            return HighlightColor.Cyan;
        }

        if (value == W.HighlightColorValues.Green)
        {
            return HighlightColor.Green;
        }

        if (value == W.HighlightColorValues.Magenta)
        {
            return HighlightColor.Magenta;
        }

        if (value == W.HighlightColorValues.Red)
        {
            return HighlightColor.Red;
        }

        if (value == W.HighlightColorValues.Yellow)
        {
            return HighlightColor.Yellow;
        }

        if (value == W.HighlightColorValues.White)
        {
            return HighlightColor.White;
        }

        if (value == W.HighlightColorValues.DarkBlue)
        {
            return HighlightColor.DarkBlue;
        }

        if (value == W.HighlightColorValues.DarkCyan)
        {
            return HighlightColor.DarkCyan;
        }

        if (value == W.HighlightColorValues.DarkGreen)
        {
            return HighlightColor.DarkGreen;
        }

        if (value == W.HighlightColorValues.DarkMagenta)
        {
            return HighlightColor.DarkMagenta;
        }

        if (value == W.HighlightColorValues.DarkRed)
        {
            return HighlightColor.DarkRed;
        }

        if (value == W.HighlightColorValues.DarkYellow)
        {
            return HighlightColor.DarkYellow;
        }

        if (value == W.HighlightColorValues.DarkGray)
        {
            return HighlightColor.DarkGray;
        }

        if (value == W.HighlightColorValues.LightGray)
        {
            return HighlightColor.LightGray;
        }

        return HighlightColor.None;
    }

    public static W.TextDirectionValues ToOpenXml(this TextDirection value) =>
        value switch
        {
            TextDirection.RotateDown => W.TextDirectionValues.TopToBottomRightToLeft,
            TextDirection.RotateUp => W.TextDirectionValues.BottomToTopLeftToRight,
            _ => W.TextDirectionValues.LefToRightTopToBottom
        };

    public static TextDirection ToTextDirection(W.TextDirectionValues value)
    {
        if (value == W.TextDirectionValues.TopToBottomRightToLeft)
        {
            return TextDirection.RotateDown;
        }

        if (value == W.TextDirectionValues.BottomToTopLeftToRight)
        {
            return TextDirection.RotateUp;
        }

        return TextDirection.Horizontal;
    }

    static Exception Unmapped<T>(T value)
        where T : struct, Enum =>
        new ArgumentOutOfRangeException(nameof(value), value, $"Unmapped {typeof(T).Name}.");
}
