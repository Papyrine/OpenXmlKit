namespace OpenXmlKit.Word;

/// <summary>
/// One level of a list definition — what its marker looks like and where it sits.
/// </summary>
public class ListLevel :
    IListLevelView
{
    internal ListLevel(int depth) =>
        Depth = depth;

    /// <summary>
    /// How deep this level is, from 0 for the outermost.
    /// </summary>
    public int Depth { get; }

    /// <summary>
    /// What the level counts in.
    /// </summary>
    public NumberFormat Format { get; set; } = NumberFormat.Decimal;

    /// <summary>
    /// The marker itself. A number placeholder is <c>%n</c> where n is the one-based level, so a
    /// second-level legal marker is <c>%1.%2.</c> and a bullet is the glyph it draws.
    /// </summary>
    public string Text { get; set; } = "";

    /// <summary>
    /// The number the level starts from.
    /// </summary>
    public int StartAt { get; set; } = 1;

    /// <summary>
    /// How the marker sits against the space reserved for it.
    /// </summary>
    public ListLevelAlignment Alignment { get; set; } = ListLevelAlignment.Left;

    /// <summary>
    /// How far the text of an item is indented.
    /// </summary>
    public Length Indent { get; set; } = Length.FromInches(0.5);

    /// <summary>
    /// How far back from the text the marker hangs.
    /// </summary>
    public Length Hanging { get; set; } = Length.FromInches(0.25);

    /// <summary>
    /// What separates the marker from the text after it.
    /// </summary>
    public ListTrailingCharacter TrailingCharacter { get; set; } = ListTrailingCharacter.Tab;

    /// <summary>
    /// Character formatting for the marker, separate from the text beside it. A bullet glyph comes
    /// from a symbol font, which is set here.
    /// </summary>
    public Font Font { get; } = new();

    // Explicit for the same reason Font.Shading is: an implicit implementation cannot narrow the
    // return type from IFontView to Font, and the write side wants the concrete one.
    IFontView IListLevelView.Font => Font;

    /// <summary>
    /// The level after which numbering restarts. Null restarts after the level above.
    /// </summary>
    public int? RestartAfterLevel { get; set; }

    internal static ListLevel Read(W.Level source)
    {
        var level = new ListLevel((int) (source.LevelIndex?.Value ?? 0));
        level.ReadFrom(source);
        return level;
    }

    internal void ReadFrom(W.Level source)
    {
        if (source.NumberingFormat?.Val is { HasValue: true } format)
        {
            Format = ToNumberFormat(format.Value);
        }

        if (source.LevelText?.Val is { HasValue: true } text)
        {
            Text = text.Value ?? "";
        }

        if (source.StartNumberingValue?.Val is { HasValue: true } start)
        {
            StartAt = start.Value;
        }

        if (source.LevelJustification?.Val is { HasValue: true } alignment)
        {
            Alignment = ToAlignment(alignment.Value);
        }

        if (source.LevelSuffix?.Val is { HasValue: true } suffix)
        {
            TrailingCharacter = ToTrailingCharacter(suffix.Value);
        }

        RestartAfterLevel = source.LevelRestart?.Val?.Value;

        if (source.PreviousParagraphProperties is { } paragraph)
        {
            var paragraphFormat = new ParagraphFormat();
            paragraphFormat.ReadFrom(CloneInto<W.ParagraphProperties>(paragraph));
            if (paragraphFormat.LeftIndent is { } indent)
            {
                Indent = indent;
            }

            if (paragraphFormat.HangingIndent is { } hanging)
            {
                Hanging = hanging;
            }
        }

        if (source.NumberingSymbolRunProperties is { } run)
        {
            Font.ReadFrom(CloneInto<W.RunProperties>(run));
        }
    }

    static T CloneInto<T>(OpenXmlElement source)
        where T : OpenXmlElement, new()
    {
        var target = new T();
        foreach (var child in source.ChildElements)
        {
            target.AppendChild(child.CloneNode(true));
        }

        return target;
    }

    internal W.Level ToOpenXml()
    {
        var level = new W.Level
        {
            LevelIndex = Depth,
            StartNumberingValue = new()
            {
                Val = StartAt
            },
            NumberingFormat = new()
            {
                Val = ToOpenXml(Format)
            },
            LevelText = new()
            {
                Val = Text
            },
            LevelJustification = new()
            {
                Val = ToOpenXml(Alignment)
            }
        };

        if (RestartAfterLevel is { } restart)
        {
            level.LevelRestart = new()
            {
                Val = restart
            };
        }

        if (TrailingCharacter != ListTrailingCharacter.Tab)
        {
            level.LevelSuffix = new()
            {
                Val = TrailingCharacter == ListTrailingCharacter.Space
                    ? W.LevelSuffixValues.Space
                    : W.LevelSuffixValues.Nothing
            };
        }

        // The marker sits in the hanging space, so the text indent and the hang are what place both
        // it and the text beside it. A missing hang puts the marker on top of the text.
        var format = new ParagraphFormat
        {
            LeftIndent = Indent,
            HangingIndent = Hanging
        };
        level.PreviousParagraphProperties = Transfer<W.PreviousParagraphProperties>(format.ToProperties());

        if (!Font.IsEmpty)
        {
            level.NumberingSymbolRunProperties = Transfer<W.NumberingSymbolRunProperties>(Font.ToProperties());
        }

        return level;
    }

    static T Transfer<T>(OpenXmlElement? source)
        where T : OpenXmlElement, new()
    {
        var target = new T();
        if (source == null)
        {
            return target;
        }

        foreach (var child in source.ChildElements.ToList())
        {
            child.Remove();
            target.AppendChild(child);
        }

        return target;
    }

    static W.NumberFormatValues ToOpenXml(NumberFormat format) =>
        format switch
        {
            NumberFormat.None => W.NumberFormatValues.None,
            NumberFormat.Bullet => W.NumberFormatValues.Bullet,
            NumberFormat.UpperRoman => W.NumberFormatValues.UpperRoman,
            NumberFormat.LowerRoman => W.NumberFormatValues.LowerRoman,
            NumberFormat.UpperLetter => W.NumberFormatValues.UpperLetter,
            NumberFormat.LowerLetter => W.NumberFormatValues.LowerLetter,
            NumberFormat.Ordinal => W.NumberFormatValues.Ordinal,
            NumberFormat.CardinalText => W.NumberFormatValues.CardinalText,
            NumberFormat.OrdinalText => W.NumberFormatValues.OrdinalText,
            NumberFormat.DecimalZero => W.NumberFormatValues.DecimalZero,
            _ => W.NumberFormatValues.Decimal
        };

    static NumberFormat ToNumberFormat(W.NumberFormatValues value)
    {
        if (value == W.NumberFormatValues.None)
        {
            return NumberFormat.None;
        }

        if (value == W.NumberFormatValues.Bullet)
        {
            return NumberFormat.Bullet;
        }

        if (value == W.NumberFormatValues.UpperRoman)
        {
            return NumberFormat.UpperRoman;
        }

        if (value == W.NumberFormatValues.LowerRoman)
        {
            return NumberFormat.LowerRoman;
        }

        if (value == W.NumberFormatValues.UpperLetter)
        {
            return NumberFormat.UpperLetter;
        }

        if (value == W.NumberFormatValues.LowerLetter)
        {
            return NumberFormat.LowerLetter;
        }

        if (value == W.NumberFormatValues.Ordinal)
        {
            return NumberFormat.Ordinal;
        }

        if (value == W.NumberFormatValues.CardinalText)
        {
            return NumberFormat.CardinalText;
        }

        if (value == W.NumberFormatValues.OrdinalText)
        {
            return NumberFormat.OrdinalText;
        }

        if (value == W.NumberFormatValues.DecimalZero)
        {
            return NumberFormat.DecimalZero;
        }

        return NumberFormat.Decimal;
    }

    static ListLevelAlignment ToAlignment(W.LevelJustificationValues value)
    {
        if (value == W.LevelJustificationValues.Center)
        {
            return ListLevelAlignment.Center;
        }

        if (value == W.LevelJustificationValues.Right)
        {
            return ListLevelAlignment.Right;
        }

        return ListLevelAlignment.Left;
    }

    static ListTrailingCharacter ToTrailingCharacter(W.LevelSuffixValues value)
    {
        if (value == W.LevelSuffixValues.Space)
        {
            return ListTrailingCharacter.Space;
        }

        if (value == W.LevelSuffixValues.Nothing)
        {
            return ListTrailingCharacter.Nothing;
        }

        return ListTrailingCharacter.Tab;
    }

    static W.LevelJustificationValues ToOpenXml(ListLevelAlignment alignment) =>
        alignment switch
        {
            ListLevelAlignment.Center => W.LevelJustificationValues.Center,
            ListLevelAlignment.Right => W.LevelJustificationValues.Right,
            _ => W.LevelJustificationValues.Left
        };
}
