namespace OpenXmlKit.Word;

/// <summary>
/// A run in a document being read.
/// </summary>
public readonly struct RunView
{
    readonly W.Run element;

    internal RunView(W.Run element) =>
        this.element = element;

    /// <summary>
    /// The run's text, with breaks and tabs rendered as the characters they stand for.
    /// </summary>
    public string Text
    {
        get
        {
            var builder = new StringBuilder();
            foreach (var child in element.ChildElements)
            {
                switch (child)
                {
                    case W.Text text:
                        builder.Append(text.Text);
                        break;
                    case W.TabChar:
                        builder.Append('\t');
                        break;
                    case W.Break:
                        builder.Append('\n');
                        break;
                    case W.SymbolChar symbol:
                        builder.Append(SymbolText(symbol));
                        break;
                }
            }

            return builder.ToString();
        }
    }

    /// <summary>
    /// The formatting written on this run. What actually applies, once styles are resolved, is
    /// <see cref="FormattingResolver.FontFor"/>.
    /// </summary>
    public IFontView Font
    {
        get
        {
            var font = new Font();
            if (element.RunProperties is { } properties)
            {
                font.ReadFrom(properties);
            }

            return font;
        }
    }

    /// <summary>
    /// Whether the run holds a drawing rather than text.
    /// </summary>
    public bool HasDrawing =>
        element.GetFirstChild<W.Drawing>() != null;

    /// <summary>
    /// The underlying OpenXML element, for anything this view does not expose.
    /// </summary>
    public W.Run ToOpenXml() =>
        element;

    internal W.Run Element => element;

    public override string ToString() =>
        Text;

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
