namespace OpenXmlKit.Word;

/// <summary>
/// The style definitions a document carries.
/// </summary>
public class Styles :
    IEnumerable<Style>
{
    readonly Document document;
    readonly Dictionary<string, Style> cache = [];

    internal Styles(Document document) =>
        this.document = document;

    W.Styles Root
    {
        get
        {
            var part = document.MainPart.StyleDefinitionsPart ??
                       document.MainPart.AddNewPart<StyleDefinitionsPart>(stylesRelationshipId);
            return part.Styles ??= new();
        }
    }

    const string stylesRelationshipId = "rStyles";

    /// <summary>
    /// The style with this id, or null.
    /// </summary>
    public Style? this[string id] => Find(id);

    /// <summary>
    /// The definition of a built-in style, if the document carries one.
    /// </summary>
    public Style? this[BuiltInStyle style] => Find(BuiltInStyleDefinitions.IdOf(style));

    public Style? Find(string id)
    {
        if (cache.TryGetValue(id, out var cached))
        {
            return cached;
        }

        var element = Root
            .Elements<W.Style>()
            .FirstOrDefault(_ => _.StyleId?.Value == id);
        if (element == null)
        {
            return null;
        }

        var style = new Style(element);
        cache[id] = style;
        return style;
    }

    public bool Contains(string id) =>
        Find(id) != null;

    public bool Contains(BuiltInStyle style) =>
        Contains(BuiltInStyleDefinitions.IdOf(style));

    /// <summary>
    /// Adds a style, or returns the one already there under this id.
    /// </summary>
    public Style Add(StyleKind kind, string id, string? name = null)
    {
        if (Find(id) is { } existing)
        {
            return existing;
        }

        var style = new Style(kind, id, name);
        Root.AppendChild(style.Element);
        cache[id] = style;
        return style;
    }

    public Style Add(StyleKind kind, string id, string? name, Action<Style> configure)
    {
        var style = Add(kind, id, name);
        configure(style);
        return style;
    }

    /// <summary>
    /// Writes Word's own definition of a built-in style into the document, unless it already has
    /// one.
    /// </summary>
    /// <remarks>
    /// Existing definitions are left untouched, so a template's customisations survive. Styles a
    /// definition depends on come with it — <see cref="BuiltInStyle.TableGrid"/> brings
    /// <see cref="BuiltInStyle.TableNormal"/>, which is where its cell padding comes from, and the
    /// heading styles bring <see cref="BuiltInStyle.Normal"/>.
    /// </remarks>
    public Style EnsureBuiltIn(BuiltInStyle style)
    {
        foreach (var dependency in BuiltInStyleDefinitions.DependenciesOf(style))
        {
            EnsureOne(dependency);
        }

        return EnsureOne(style);
    }

    /// <summary>
    /// Writes Word's own definitions of several built-in styles.
    /// </summary>
    public Styles EnsureBuiltIn(params BuiltInStyle[] styles)
    {
        foreach (var style in styles)
        {
            EnsureBuiltIn(style);
        }

        return this;
    }

    Style EnsureOne(BuiltInStyle style)
    {
        var id = BuiltInStyleDefinitions.IdOf(style);
        if (Find(id) is { } existing)
        {
            return existing;
        }

        var element = BuiltInStyleDefinitions.Build(style);
        Root.AppendChild(element);
        var wrapper = new Style(element);
        cache[id] = wrapper;
        return wrapper;
    }

    /// <summary>
    /// The document defaults — the formatting everything inherits before any style applies.
    /// </summary>
    public void SetDefaults(Action<Font>? font = null, Action<ParagraphFormat>? paragraph = null)
    {
        var defaults = Root.GetFirstChild<W.DocDefaults>();
        if (defaults == null)
        {
            defaults = new();
            // docDefaults leads the styles part, ahead of every style definition.
            Root.PrependChild(defaults);
        }

        if (font != null)
        {
            var model = new Font();
            font(model);
            var runDefault = new W.RunPropertiesDefault();
            var runProperties = new W.RunPropertiesBaseStyle();
            if (model.ToProperties() is { } source)
            {
                foreach (var child in source.ChildElements.ToList())
                {
                    child.Remove();
                    runProperties.AppendChild(child);
                }
            }

            runDefault.RunPropertiesBaseStyle = runProperties;
            defaults.RunPropertiesDefault = runDefault;
        }

        if (paragraph == null)
        {
            return;
        }

        var paragraphModel = new ParagraphFormat();
        paragraph(paragraphModel);
        var paragraphDefault = new W.ParagraphPropertiesDefault();
        var paragraphProperties = new W.ParagraphPropertiesBaseStyle();
        if (paragraphModel.ToProperties() is { } paragraphSource)
        {
            foreach (var child in paragraphSource.ChildElements.ToList())
            {
                child.Remove();
                paragraphProperties.AppendChild(child);
            }
        }

        paragraphDefault.ParagraphPropertiesBaseStyle = paragraphProperties;
        defaults.ParagraphPropertiesDefault = paragraphDefault;
    }

    public IEnumerator<Style> GetEnumerator()
    {
        foreach (var element in Root.Elements<W.Style>())
        {
            var id = element.StyleId?.Value ?? "";
            if (cache.TryGetValue(id, out var cached))
            {
                yield return cached;
                continue;
            }

            var style = new Style(element);
            cache[id] = style;
            yield return style;
        }
    }

    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();

    internal void Save()
    {
        foreach (var style in cache.Values)
        {
            style.Flush();
        }

        document.MainPart.StyleDefinitionsPart?.Styles?.Save();
    }
}
