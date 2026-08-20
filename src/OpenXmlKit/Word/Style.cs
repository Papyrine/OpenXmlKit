namespace OpenXmlKit.Word;

/// <summary>
/// A named style definition.
/// </summary>
public class Style
{
    readonly W.Style element;
    Font? font;
    ParagraphFormat? paragraphFormat;
    TableFormat? tableFormat;

    internal Style(W.Style element)
    {
        this.element = element;
        if (element.StyleRunProperties is { } runProperties)
        {
            font = new();
            font.ReadFrom(ToRunProperties(runProperties));
        }

        if (element.StyleParagraphProperties is { } paragraphProperties)
        {
            paragraphFormat = new();
            paragraphFormat.ReadFrom(ToParagraphProperties(paragraphProperties));
        }
    }

    internal Style(StyleKind kind, string id, string? name) =>
        element = new()
        {
            Type = kind.ToOpenXml(),
            StyleId = id,
            StyleName = new()
            {
                Val = name ?? id
            }
        };

    public string Id
    {
        get => element.StyleId?.Value ?? "";
        set => element.StyleId = value;
    }

    public string Name
    {
        get => element.StyleName?.Val?.Value ?? Id;
        set => element.StyleName = new()
        {
            Val = value
        };
    }

    public StyleKind Kind =>
        element.Type is { HasValue: true } type ? Map.ToStyleKind(type.Value) : StyleKind.Paragraph;

    /// <summary>
    /// The style this one inherits from.
    /// </summary>
    public string? BasedOn
    {
        get => element.BasedOn?.Val?.Value;
        set => element.BasedOn = value == null
            ? null
            : new W.BasedOn
            {
                Val = value
            };
    }

    /// <summary>
    /// The style applied to the paragraph typed after one in this style — how a heading is
    /// followed by body text.
    /// </summary>
    public string? NextStyle
    {
        get => element.NextParagraphStyle?.Val?.Value;
        set => element.NextParagraphStyle = value == null
            ? null
            : new W.NextParagraphStyle
            {
                Val = value
            };
    }

    /// <summary>
    /// The character style paired with this paragraph style.
    /// </summary>
    public string? LinkedStyle
    {
        get => element.LinkedStyle?.Val?.Value;
        set => element.LinkedStyle = value == null
            ? null
            : new W.LinkedStyle
            {
                Val = value
            };
    }

    /// <summary>
    /// Where the style sorts in Word's gallery. Lower comes first.
    /// </summary>
    public int? Priority
    {
        get => element.UIPriority?.Val?.Value;
        set => element.UIPriority = value == null
            ? null
            : new W.UIPriority
            {
                Val = value
            };
    }

    /// <summary>
    /// Shows the style in Word's gallery.
    /// </summary>
    public bool IsQuickStyle
    {
        get => element.PrimaryStyle != null;
        set => element.PrimaryStyle = value ? new W.PrimaryStyle() : null;
    }

    public bool SemiHidden
    {
        get => element.SemiHidden != null;
        set => element.SemiHidden = value ? new W.SemiHidden() : null;
    }

    public bool UnhideWhenUsed
    {
        get => element.UnhideWhenUsed != null;
        set => element.UnhideWhenUsed = value ? new W.UnhideWhenUsed() : null;
    }

    /// <summary>
    /// Whether this is the default style of its kind, applied to content that names none.
    /// </summary>
    public bool IsDefault
    {
        get => element.Default?.Value == true;
        set => element.Default = value;
    }

    /// <summary>
    /// Character formatting the style contributes.
    /// </summary>
    public Font Font => font ??= new();

    /// <summary>
    /// Paragraph formatting the style contributes.
    /// </summary>
    public ParagraphFormat ParagraphFormat => paragraphFormat ??= new();

    /// <summary>
    /// Table formatting the style contributes. Table styles only.
    /// </summary>
    public TableFormat TableFormat => tableFormat ??= new();

    public W.Style ToOpenXml()
    {
        Flush();
        return element;
    }

    internal W.Style Element
    {
        get
        {
            Flush();
            return element;
        }
    }

    // The three property blocks live in CT_Style's sequence after name/basedOn/next/uiPriority.
    // Building each in its own element type and assigning through the typed property is what puts
    // them in that order, and what orders their own children within it.
    internal void Flush()
    {
        if (font is { IsEmpty: false })
        {
            var properties = new W.StyleRunProperties();
            var source = font.ToProperties();
            if (source != null)
            {
                MoveChildren(source, properties);
            }

            element.StyleRunProperties = properties;
        }

        if (paragraphFormat is { IsEmpty: false })
        {
            var properties = new W.StyleParagraphProperties();
            var source = paragraphFormat.ToProperties();
            if (source != null)
            {
                MoveChildren(source, properties);
            }

            element.StyleParagraphProperties = properties;
        }

        if (tableFormat is { IsEmpty: false })
        {
            element.GetFirstChild<W.StyleTableProperties>()?.Remove();
            element.Append(tableFormat.ToStyleProperties());
        }
    }

    // The style-scoped property elements carry the same children in the same order as the
    // content-scoped ones, so the ordering the SDK applied when building the source is reused
    // rather than restated.
    static void MoveChildren(OpenXmlElement source, OpenXmlElement target)
    {
        foreach (var child in source.ChildElements.ToList())
        {
            child.Remove();
            target.AppendChild(child);
        }
    }

    static W.RunProperties ToRunProperties(W.StyleRunProperties source)
    {
        var properties = new W.RunProperties();
        foreach (var child in source.ChildElements)
        {
            properties.AppendChild(child.CloneNode(true));
        }

        return properties;
    }

    static W.ParagraphProperties ToParagraphProperties(W.StyleParagraphProperties source)
    {
        var properties = new W.ParagraphProperties();
        foreach (var child in source.ChildElements)
        {
            properties.AppendChild(child.CloneNode(true));
        }

        return properties;
    }
}
