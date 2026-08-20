namespace OpenXmlKit.Word;

/// <summary>
/// The style definitions of a document being read.
/// </summary>
public sealed class StylesView :
    IEnumerable<StyleView>
{
    readonly MainDocumentPart main;

    internal StylesView(MainDocumentPart main) =>
        this.main = main;

    W.Styles? Root => main.StyleDefinitionsPart?.Styles;

    /// <summary>
    /// The style with this id, or null.
    /// </summary>
    public StyleView? this[string id] => Find(id);

    /// <summary>
    /// The definition of a built-in style, if the document carries one.
    /// </summary>
    public StyleView? this[BuiltInStyle style] => Find(BuiltInStyleDefinitions.IdOf(style));

    /// <summary>
    /// The style with the given id, or null when the document has none.
    /// </summary>
    public StyleView? Find(string id)
    {
        foreach (var element in Root?.Elements<W.Style>() ?? [])
        {
            if (element.StyleId?.Value == id)
            {
                return new(element);
            }
        }

        return null;
    }

    /// <summary>
    /// Whether the document defines a style with the given id.
    /// </summary>
    public bool Contains(string id) =>
        Find(id) != null;

    /// <summary>
    /// The document defaults — the formatting everything inherits before any style applies.
    /// </summary>
    public IFontView? DefaultFont
    {
        get
        {
            if (Root?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle is not { } source)
            {
                return null;
            }

            var properties = new W.RunProperties();
            foreach (var child in source.ChildElements)
            {
                properties.AppendChild(child.CloneNode(true));
            }

            var font = new Font();
            font.ReadFrom(properties);
            return font;
        }
    }

    /// <summary>
    /// The paragraph formatting in the document defaults, which every style starts from.
    /// </summary>
    public IParagraphFormatView? DefaultParagraphFormat
    {
        get
        {
            if (Root?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle is not { } source)
            {
                return null;
            }

            var properties = new W.ParagraphProperties();
            foreach (var child in source.ChildElements)
            {
                properties.AppendChild(child.CloneNode(true));
            }

            var format = new ParagraphFormat();
            format.ReadFrom(properties);
            return format;
        }
    }

    /// <summary>
    /// Every style the document defines.
    /// </summary>
    public IEnumerator<StyleView> GetEnumerator()
    {
        foreach (var element in Root?.Elements<W.Style>() ?? [])
        {
            yield return new(element);
        }
    }

    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();
}
