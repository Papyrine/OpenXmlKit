namespace OpenXmlKit.Word;

/// <summary>
/// A style definition in a document being read.
/// </summary>
public readonly struct StyleView
{
    readonly W.Style element;

    internal StyleView(W.Style element) =>
        this.element = element;

    public string Id => element.StyleId?.Value ?? "";

    public string Name => element.StyleName?.Val?.Value ?? Id;

    public StyleKind Kind =>
        element.Type is { HasValue: true } type ? Map.ToStyleKind(type.Value) : StyleKind.Paragraph;

    /// <summary>
    /// The style this one inherits from.
    /// </summary>
    public string? BasedOn => element.BasedOn?.Val?.Value;

    public string? NextStyle => element.NextParagraphStyle?.Val?.Value;

    public string? LinkedStyle => element.LinkedStyle?.Val?.Value;

    public int? Priority => element.UIPriority?.Val?.Value;

    public bool IsDefault => element.Default?.Value == true;

    /// <summary>
    /// The character formatting this style contributes, before its inheritance chain is walked.
    /// </summary>
    public IFontView Font
    {
        get
        {
            var font = new Font();
            if (element.StyleRunProperties is { } source)
            {
                var properties = new W.RunProperties();
                foreach (var child in source.ChildElements)
                {
                    properties.AppendChild(child.CloneNode(true));
                }

                font.ReadFrom(properties);
            }

            return font;
        }
    }

    /// <summary>
    /// The paragraph formatting this style contributes, before its inheritance chain is walked.
    /// </summary>
    public IParagraphFormatView ParagraphFormat
    {
        get
        {
            var format = new ParagraphFormat();
            if (element.StyleParagraphProperties is { } source)
            {
                var properties = new W.ParagraphProperties();
                foreach (var child in source.ChildElements)
                {
                    properties.AppendChild(child.CloneNode(true));
                }

                format.ReadFrom(properties);
            }

            return format;
        }
    }

    /// <summary>
    /// The underlying OpenXML element, for anything this view does not expose.
    /// </summary>
    public W.Style ToOpenXml() =>
        element;

    public override string ToString() =>
        Id;
}
