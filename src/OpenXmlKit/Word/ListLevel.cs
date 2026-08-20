namespace OpenXmlKit.Word;

/// <summary>
/// One level of a list definition — what its marker looks like and where it sits.
/// </summary>
public class ListLevel
{
    internal ListLevel(int depth) =>
        Depth = depth;

    /// <summary>
    /// How deep this level is, from 0 for the outermost.
    /// </summary>
    public int Depth { get; }

    public NumberFormat Format { get; set; } = NumberFormat.Decimal;

    /// <summary>
    /// The marker itself. A number placeholder is <c>%n</c> where n is the one-based level, so a
    /// second-level legal marker is <c>%1.%2.</c> and a bullet is the glyph it draws.
    /// </summary>
    public string Text { get; set; } = "";

    public int StartAt { get; set; } = 1;

    public ListLevelAlignment Alignment { get; set; } = ListLevelAlignment.Left;

    /// <summary>
    /// How far the text of an item is indented.
    /// </summary>
    public Length Indent { get; set; } = Length.FromInches(0.5);

    /// <summary>
    /// How far back from the text the marker hangs.
    /// </summary>
    public Length Hanging { get; set; } = Length.FromInches(0.25);

    public ListTrailingCharacter TrailingCharacter { get; set; } = ListTrailingCharacter.Tab;

    /// <summary>
    /// Character formatting for the marker, separate from the text beside it. A bullet glyph comes
    /// from a symbol font, which is set here.
    /// </summary>
    public Font Font { get; } = new();

    /// <summary>
    /// The level after which numbering restarts. Null restarts after the level above.
    /// </summary>
    public int? RestartAfterLevel { get; set; }

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

    static W.LevelJustificationValues ToOpenXml(ListLevelAlignment alignment) =>
        alignment switch
        {
            ListLevelAlignment.Center => W.LevelJustificationValues.Center,
            ListLevelAlignment.Right => W.LevelJustificationValues.Right,
            _ => W.LevelJustificationValues.Left
        };
}
