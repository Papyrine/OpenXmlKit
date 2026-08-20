namespace OpenXmlKit.Word;

/// <summary>
/// Character formatting: everything a run property can say about how text looks.
/// </summary>
/// <remarks>
/// One type, mounted everywhere character formatting appears — on a run, on a style, on the
/// builder cursor, on a list level's number. Aspose does the same, and the consistency is worth
/// more than a family of near-identical types would be.
/// <para>
/// Every property is "unstated" by default: a null, or a <see cref="Toggle.Inherit"/>. An
/// untouched font emits no run properties at all, so the style hierarchy is left to resolve. That
/// is what makes it safe to hand a shared font object down through a builder and only state the
/// one thing that differs.
/// </para>
/// </remarks>
public partial class Font
{
    /// <summary>
    /// The typeface, applied to every script. Set <see cref="NameAscii"/> and the rest directly to
    /// vary it by script.
    /// </summary>
    public string? Name
    {
        get => NameAscii ?? NameHighAnsi ?? NameComplexScript ?? NameEastAsia;
        set
        {
            NameAscii = value;
            NameHighAnsi = value;
        }
    }

    public string? NameAscii { get; set; }
    public string? NameHighAnsi { get; set; }
    public string? NameComplexScript { get; set; }
    public string? NameEastAsia { get; set; }

    /// <summary>
    /// Point size. Word stores it in half-points, which is where <c>sz=24</c> meaning 12pt comes
    /// from; <see cref="Length"/> keeps that off the call site.
    /// </summary>
    public Length? Size { get; set; }

    /// <summary>
    /// Point size for complex-script runs. Falls back to <see cref="Size"/> when unset.
    /// </summary>
    public Length? SizeComplexScript { get; set; }

    public Toggle Bold { get; set; }
    public Toggle BoldComplexScript { get; set; }
    public Toggle Italic { get; set; }
    public Toggle ItalicComplexScript { get; set; }

    public UnderlineStyle? Underline { get; set; }
    public Color? UnderlineColor { get; set; }

    public Toggle Strike { get; set; }
    public Toggle DoubleStrike { get; set; }
    public Toggle SmallCaps { get; set; }
    public Toggle AllCaps { get; set; }
    public Toggle Outline { get; set; }
    public Toggle Shadow { get; set; }
    public Toggle Emboss { get; set; }
    public Toggle Imprint { get; set; }

    /// <summary>
    /// Hidden text — present in the document, not shown unless Word is set to reveal it.
    /// </summary>
    public Toggle Hidden { get; set; }

    public Toggle NoProof { get; set; }
    public Toggle RightToLeft { get; set; }

    public Color? Color { get; set; }

    /// <summary>
    /// The highlighter pen, which has a fixed palette. For an arbitrary colour behind text, use
    /// <see cref="Shading"/>.
    /// </summary>
    public HighlightColor? Highlight { get; set; }

    public Shading Shading { get; } = new();

    /// <summary>
    /// A border drawn around the run.
    /// </summary>
    public Border Border { get; } = new();

    public VerticalTextPosition? VerticalPosition { get; set; }

    /// <summary>
    /// Extra space between characters. Negative values tighten.
    /// </summary>
    public Length? CharacterSpacing { get; set; }

    /// <summary>
    /// Horizontal stretch, as a percentage. 100 is unstretched.
    /// </summary>
    public int? Scale { get; set; }

    /// <summary>
    /// Raises text above the baseline, or lowers it when negative. Unlike
    /// <see cref="VerticalPosition"/>, this does not shrink the text.
    /// </summary>
    public Length? Position { get; set; }

    /// <summary>
    /// The id of a character style to apply. Direct formatting on this font still wins over it.
    /// </summary>
    public string? StyleId { get; set; }

    /// <summary>
    /// Whether anything at all is stated. An empty font writes no run properties.
    /// </summary>
    public bool IsEmpty =>
        NameAscii == null &&
        NameHighAnsi == null &&
        NameComplexScript == null &&
        NameEastAsia == null &&
        Size == null &&
        SizeComplexScript == null &&
        !Bold.IsSet &&
        !BoldComplexScript.IsSet &&
        !Italic.IsSet &&
        !ItalicComplexScript.IsSet &&
        Underline == null &&
        UnderlineColor == null &&
        !Strike.IsSet &&
        !DoubleStrike.IsSet &&
        !SmallCaps.IsSet &&
        !AllCaps.IsSet &&
        !Outline.IsSet &&
        !Shadow.IsSet &&
        !Emboss.IsSet &&
        !Imprint.IsSet &&
        !Hidden.IsSet &&
        !NoProof.IsSet &&
        !RightToLeft.IsSet &&
        Color == null &&
        Highlight == null &&
        Shading.IsEmpty &&
        Border.IsEmpty &&
        VerticalPosition == null &&
        CharacterSpacing == null &&
        Scale == null &&
        Position == null &&
        StyleId == null;

    public Font Clone()
    {
        var clone = new Font();
        clone.CopyFrom(this);
        return clone;
    }

    public void CopyFrom(Font other)
    {
        NameAscii = other.NameAscii;
        NameHighAnsi = other.NameHighAnsi;
        NameComplexScript = other.NameComplexScript;
        NameEastAsia = other.NameEastAsia;
        Size = other.Size;
        SizeComplexScript = other.SizeComplexScript;
        Bold = other.Bold;
        BoldComplexScript = other.BoldComplexScript;
        Italic = other.Italic;
        ItalicComplexScript = other.ItalicComplexScript;
        Underline = other.Underline;
        UnderlineColor = other.UnderlineColor;
        Strike = other.Strike;
        DoubleStrike = other.DoubleStrike;
        SmallCaps = other.SmallCaps;
        AllCaps = other.AllCaps;
        Outline = other.Outline;
        Shadow = other.Shadow;
        Emboss = other.Emboss;
        Imprint = other.Imprint;
        Hidden = other.Hidden;
        NoProof = other.NoProof;
        RightToLeft = other.RightToLeft;
        Color = other.Color;
        Highlight = other.Highlight;
        Shading.CopyFrom(other.Shading);
        Border.CopyFrom(other.Border);
        VerticalPosition = other.VerticalPosition;
        CharacterSpacing = other.CharacterSpacing;
        Scale = other.Scale;
        Position = other.Position;
        StyleId = other.StyleId;
    }

    /// <summary>
    /// Returns everything to unstated, so the style hierarchy resolves the lot.
    /// </summary>
    public void Clear() =>
        CopyFrom(new());

    /// <summary>
    /// Builds the run properties, or null when nothing is stated.
    /// </summary>
    /// <remarks>
    /// Every child is assigned through a typed property rather than appended, which is what places
    /// it at its position in the CT_RPr sequence. Word treats an out-of-order rPr as a corrupt
    /// document, so the ordering is not a matter of tidiness — see SchemaOrderTests.
    /// </remarks>
    internal W.RunProperties? ToProperties()
    {
        if (IsEmpty)
        {
            return null;
        }

        var properties = new W.RunProperties();
        ApplyTo(properties);
        return properties;
    }

    internal void ApplyTo(W.RunProperties properties)
    {
        if (StyleId != null)
        {
            properties.RunStyle = new()
            {
                Val = StyleId
            };
        }

        if (NameAscii != null ||
            NameHighAnsi != null ||
            NameComplexScript != null ||
            NameEastAsia != null)
        {
            var fonts = new W.RunFonts();
            if (NameAscii != null)
            {
                fonts.Ascii = NameAscii;
            }

            if (NameHighAnsi != null)
            {
                fonts.HighAnsi = NameHighAnsi;
            }

            if (NameComplexScript != null)
            {
                fonts.ComplexScript = NameComplexScript;
            }

            if (NameEastAsia != null)
            {
                fonts.EastAsia = NameEastAsia;
            }

            properties.RunFonts = fonts;
        }

        properties.Bold = Toggles.OnOff<W.Bold>(Bold);
        properties.BoldComplexScript = Toggles.OnOff<W.BoldComplexScript>(BoldComplexScript);
        properties.Italic = Toggles.OnOff<W.Italic>(Italic);
        properties.ItalicComplexScript = Toggles.OnOff<W.ItalicComplexScript>(ItalicComplexScript);
        properties.Caps = Toggles.OnOff<W.Caps>(AllCaps);
        properties.SmallCaps = Toggles.OnOff<W.SmallCaps>(SmallCaps);
        properties.Strike = Toggles.OnOff<W.Strike>(Strike);
        properties.DoubleStrike = Toggles.OnOff<W.DoubleStrike>(DoubleStrike);
        properties.Outline = Toggles.OnOff<W.Outline>(Outline);
        properties.Shadow = Toggles.OnOff<W.Shadow>(Shadow);
        properties.Emboss = Toggles.OnOff<W.Emboss>(Emboss);
        properties.Imprint = Toggles.OnOff<W.Imprint>(Imprint);
        properties.NoProof = Toggles.OnOff<W.NoProof>(NoProof);
        properties.Vanish = Toggles.OnOff<W.Vanish>(Hidden);
        properties.RightToLeftText = Toggles.OnOff<W.RightToLeftText>(RightToLeft);

        if (Color is { } color)
        {
            var element = new W.Color
            {
                Val = color.Value
            };
            if (color.IsTheme)
            {
                element.ThemeColor = color.Theme.ToOpenXml();
            }

            properties.Color = element;
        }

        if (CharacterSpacing is { } spacing)
        {
            properties.Spacing = new()
            {
                Val = spacing.Twips
            };
        }

        if (Scale is { } scale)
        {
            properties.CharacterScale = new()
            {
                Val = scale
            };
        }

        if (Size is { } size)
        {
            properties.FontSize = new()
            {
                Val = size.HalfPoints.ToString(CultureInfo.InvariantCulture)
            };
        }

        if (SizeComplexScript is { } sizeComplex)
        {
            properties.FontSizeComplexScript = new()
            {
                Val = sizeComplex.HalfPoints.ToString(CultureInfo.InvariantCulture)
            };
        }

        if (Highlight is { } highlight)
        {
            properties.Highlight = new()
            {
                Val = highlight.ToOpenXml()
            };
        }

        if (Underline is { } underline)
        {
            var element = new W.Underline
            {
                Val = underline.ToOpenXml()
            };
            if (UnderlineColor is { } underlineColor)
            {
                element.Color = underlineColor.Value;
            }

            properties.Underline = element;
        }

        if (!Border.IsEmpty)
        {
            var border = new W.Border();
            Border.ApplyTo(border);
            properties.Border = border;
        }

        if (Shading.ToOpenXml() is { } shading)
        {
            properties.Shading = shading;
        }

        if (Position is { } position)
        {
            properties.Position = new()
            {
                Val = position.HalfPoints.ToString(CultureInfo.InvariantCulture)
            };
        }

        if (VerticalPosition is { } verticalPosition)
        {
            properties.VerticalTextAlignment = new()
            {
                Val = verticalPosition.ToOpenXml()
            };
        }
    }

    // A stated-off toggle still emits its element, carrying w:val="0". That is the whole point of
    // the three-state model: it is the only way to cancel formatting a style turned on. An
    // inherited toggle returns null, and assigning null to a typed property removes the child.

    internal static Font? Read(W.RunProperties? properties)
    {
        if (properties == null)
        {
            return null;
        }

        var font = new Font();
        font.ReadFrom(properties);
        return font;
    }

    internal void ReadFrom(W.RunProperties properties)
    {
        StyleId = properties.RunStyle?.Val?.Value;

        if (properties.RunFonts is { } fonts)
        {
            NameAscii = fonts.Ascii?.Value;
            NameHighAnsi = fonts.HighAnsi?.Value;
            NameComplexScript = fonts.ComplexScript?.Value;
            NameEastAsia = fonts.EastAsia?.Value;
        }

        Bold = Toggles.Read(properties.Bold);
        BoldComplexScript = Toggles.Read(properties.BoldComplexScript);
        Italic = Toggles.Read(properties.Italic);
        ItalicComplexScript = Toggles.Read(properties.ItalicComplexScript);
        AllCaps = Toggles.Read(properties.Caps);
        SmallCaps = Toggles.Read(properties.SmallCaps);
        Strike = Toggles.Read(properties.Strike);
        DoubleStrike = Toggles.Read(properties.DoubleStrike);
        Outline = Toggles.Read(properties.Outline);
        Shadow = Toggles.Read(properties.Shadow);
        Emboss = Toggles.Read(properties.Emboss);
        Imprint = Toggles.Read(properties.Imprint);
        NoProof = Toggles.Read(properties.NoProof);
        Hidden = Toggles.Read(properties.Vanish);
        RightToLeft = Toggles.Read(properties.RightToLeftText);

        if (properties.Color?.Val is { HasValue: true } color &&
            Word.Color.TryParse(color.Value, out var parsed))
        {
            Color = parsed;
        }

        if (properties.Spacing?.Val is { HasValue: true } spacing)
        {
            CharacterSpacing = Length.FromTwips(spacing.Value);
        }

        if (properties.CharacterScale?.Val is { HasValue: true } scale)
        {
            Scale = (int) scale.Value;
        }

        Size = ReadHalfPoints(properties.FontSize?.Val?.Value);
        SizeComplexScript = ReadHalfPoints(properties.FontSizeComplexScript?.Val?.Value);
        Position = ReadHalfPoints(properties.Position?.Val?.Value);

        if (properties.Highlight?.Val is { HasValue: true } highlight)
        {
            Highlight = Map.ToHighlight(highlight.Value);
        }

        if (properties.Underline is { } underline)
        {
            if (underline.Val is { HasValue: true } style)
            {
                Underline = Map.ToUnderline(style.Value);
            }

            if (underline.Color is { HasValue: true } underlineColor &&
                Word.Color.TryParse(underlineColor.Value, out var parsedUnderline))
            {
                UnderlineColor = parsedUnderline;
            }
        }

        if (properties.Border is { } border)
        {
            Border.ReadFrom(border);
        }

        Shading.ReadFrom(properties.Shading);

        if (properties.VerticalTextAlignment?.Val is { HasValue: true } verticalPosition)
        {
            VerticalPosition = Map.ToVerticalTextPosition(verticalPosition.Value);
        }
    }

    static Length? ReadHalfPoints(string? value)
    {
        if (value == null ||
            !double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var halfPoints))
        {
            return null;
        }

        return Length.FromHalfPoints(halfPoints);
    }

}
