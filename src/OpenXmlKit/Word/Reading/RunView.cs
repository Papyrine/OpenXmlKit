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
                    case W.CarriageReturn:
                        builder.Append('\n');
                        break;
                    case W.PositionalTab:
                        builder.Append('\t');
                        break;
                    // A non-breaking hyphen and a soft hyphen are characters Word stores as
                    // elements rather than as text, so a reader that only looks at w:t drops them
                    // and silently joins the words either side.
                    case W.NoBreakHyphen:
                        builder.Append('\u2011');
                        break;
                    case W.SoftHyphen:
                        builder.Append('\u00AD');
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
    /// The picture in this run, or null when it holds none.
    /// </summary>
    public ImageView? Image =>
        element.GetFirstChild<W.Drawing>() is { } drawing ? new ImageView(drawing) : null;

    /// <summary>
    /// The id of the footnote this run is the reference mark for, or null when it is not one.
    /// </summary>
    public int? FootnoteReference =>
        element.GetFirstChild<W.FootnoteReference>()?.Id?.Value is { } id ? (int) id : null;

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
