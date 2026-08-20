namespace OpenXmlKit.Word;

/// <summary>
/// A stretch of text sharing one set of character formatting.
/// </summary>
public class Run
{
    readonly W.Run element;
    Font? font;

    public Run() =>
        element = new();

    /// <summary>
    /// Character formatting for this run.
    /// </summary>
    public Font Font => font ??= new();

    /// <summary>
    /// Appends text.
    /// </summary>
    /// <remarks>
    /// Newlines become line breaks and tabs become tab characters, because Word ignores both when
    /// they appear inside a text element — a value carrying its own line breaks would otherwise
    /// silently run together. Leading and trailing whitespace is preserved rather than collapsed.
    /// </remarks>
    public Run Append(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return this;
        }

        var normalized = text!.Replace("\r\n", "\n").Replace('\r', '\n');
        var start = 0;
        for (var index = 0; index < normalized.Length; index++)
        {
            var character = normalized[index];
            if (character is not ('\n' or '\t'))
            {
                continue;
            }

            AppendText(normalized[start..index]);
            element.Append(character == '\n' ? new W.Break() : new W.TabChar());
            start = index + 1;
        }

        AppendText(normalized[start..]);
        return this;
    }

    void AppendText(string value)
    {
        if (value.Length == 0)
        {
            return;
        }

        element.Append(
            new W.Text(value)
            {
                Space = SpaceProcessingModeValues.Preserve
            });
    }

    /// <summary>
    /// Appends a break.
    /// </summary>
    public Run AppendBreak(BreakKind kind = BreakKind.Line)
    {
        var brk = new W.Break();
        if (kind == BreakKind.Page)
        {
            brk.Type = W.BreakValues.Page;
        }
        else if (kind == BreakKind.Column)
        {
            brk.Type = W.BreakValues.Column;
        }

        element.Append(brk);
        return this;
    }

    public Run AppendTab()
    {
        element.Append(new W.TabChar());
        return this;
    }

    /// <summary>
    /// Appends an element this library does not model — a drawing, a field character, a note
    /// reference.
    /// </summary>
    public Run AppendElement(OpenXmlElement child)
    {
        element.AppendChild(child);
        return this;
    }

    /// <summary>
    /// Appends a glyph from a symbol font, addressed by its character code.
    /// </summary>
    /// <remarks>
    /// This is how a checkbox is drawn without a list: Wingdings F0FE is a ticked box and F06F an
    /// empty one. The glyph comes from the named font regardless of the run's own typeface.
    /// </remarks>
    public Run AppendSymbol(string fontName, char character)
    {
        element.Append(
            new W.SymbolChar
            {
                Font = fontName,
                Char = ((int) character).ToString("X4", CultureInfo.InvariantCulture)
            });
        return this;
    }

    public Run Bold(Toggle value = default)
    {
        Font.Bold = value.IsSet ? value : Toggle.On;
        return this;
    }

    public Run Italic(Toggle value = default)
    {
        Font.Italic = value.IsSet ? value : Toggle.On;
        return this;
    }

    public Run Underline(UnderlineStyle style = UnderlineStyle.Single)
    {
        Font.Underline = style;
        return this;
    }

    public Run Color(Color value)
    {
        Font.Color = value;
        return this;
    }

    public Run Size(Length value)
    {
        Font.Size = value;
        return this;
    }

    /// <summary>
    /// Applies a character style by id.
    /// </summary>
    public Run Style(string styleId)
    {
        Font.StyleId = styleId;
        return this;
    }

    /// <summary>
    /// Configures the character formatting.
    /// </summary>
    public Run Formatting(Action<Font> configure)
    {
        configure(Font);
        return this;
    }

    /// <summary>
    /// The underlying OpenXML element, with any pending formatting applied.
    /// </summary>
    /// <remarks>
    /// The escape hatch. Anything this library does not model can be reached by working on the
    /// element directly, and a raw element built elsewhere can be handed back in — which is what
    /// keeps a partial migration onto this library possible.
    /// </remarks>
    public W.Run ToOpenXml()
    {
        Flush();
        return element;
    }

    internal W.Run Element => element;

    // Rebuilds rPr from the format object wholesale, so flushing twice produces the same result as
    // flushing once and a caller can keep editing after reading the element.
    internal void Flush()
    {
        if (font == null)
        {
            return;
        }

        element.RunProperties = font.ToProperties();
    }

    static string SymbolText(W.SymbolChar symbol)
    {
        if (symbol.Char is not { HasValue: true } code ||
            !int.TryParse(code.Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var value))
        {
            return "";
        }

        return char.ConvertFromUtf32(value);
    }
}
