namespace OpenXmlKit.Word;

/// <summary>
/// A style definition in a document being read.
/// </summary>
public readonly struct StyleView
{
    readonly W.Style element;

    internal StyleView(W.Style element) =>
        this.element = element;

    /// <summary>
    /// The id content refers to.
    /// </summary>
    public string Id => element.StyleId?.Value ?? "";

    /// <summary>
    /// The name Word shows, falling back to the id when the style has none.
    /// </summary>
    public string Name => element.StyleName?.Val?.Value ?? Id;

    /// <summary>
    /// What the style can be applied to.
    /// </summary>
    public StyleKind Kind =>
        element.Type is { HasValue: true } type ? Map.ToStyleKind(type.Value) : StyleKind.Paragraph;

    /// <summary>
    /// The style this one inherits from.
    /// </summary>
    public string? BasedOn => element.BasedOn?.Val?.Value;

    /// <summary>
    /// The style applied to the paragraph typed after one in this style.
    /// </summary>
    public string? NextStyle => element.NextParagraphStyle?.Val?.Value;

    /// <summary>
    /// The character style paired with this paragraph style.
    /// </summary>
    public string? LinkedStyle => element.LinkedStyle?.Val?.Value;

    /// <summary>
    /// Where the style sorts in the gallery. Lower comes first.
    /// </summary>
    public int? Priority => element.UIPriority?.Val?.Value;

    /// <summary>
    /// Whether this is the default style of its kind.
    /// </summary>
    public bool IsDefault => element.Default?.Value == true;

    /// <summary>
    /// Whether the style shows in Word's gallery.
    /// </summary>
    public bool IsQuickStyle => element.PrimaryStyle != null;

    /// <summary>
    /// Whether the style is hidden from the gallery until used.
    /// </summary>
    public bool SemiHidden => element.SemiHidden != null;

    /// <summary>
    /// Whether using the style reveals it.
    /// </summary>
    public bool UnhideWhenUsed => element.UnhideWhenUsed != null;

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
    /// The table formatting this style contributes, before its inheritance chain is walked. Empty
    /// for a style that is not a table style.
    /// </summary>
    public ITableFormatView TableFormat
    {
        get
        {
            var format = new TableFormat();
            format.ReadFrom(CloneInto<W.TableProperties>(element.GetFirstChild<W.StyleTableProperties>()));
            return format;
        }
    }

    /// <summary>
    /// The formatting the style applies to the individual parts of a table — the header row, the
    /// banding, the corner cells.
    /// </summary>
    /// <remarks>
    /// A table style is mostly these. Reading only <see cref="TableFormat"/> sees the whole-table
    /// block and none of what makes the style recognisable.
    /// </remarks>
    public IEnumerable<ITableStyleConditionalView> ConditionalFormats
    {
        get
        {
            foreach (var source in element.Elements<W.TableStyleProperties>())
            {
                var area = source.Type is { HasValue: true } type
                    ? Map.ToTableStyleArea(type.Value)
                    : TableStyleArea.WholeTable;
                var conditional = new TableStyleConditional(area);
                conditional.Font.ReadFrom(CloneInto<W.RunProperties>(source.RunPropertiesBaseStyle));
                conditional.ParagraphFormat.ReadFrom(CloneInto<W.ParagraphProperties>(source.StyleParagraphProperties));
                conditional.TableFormat.ReadFrom(CloneInto<W.TableProperties>(source.TableStyleConditionalFormattingTableProperties));
                conditional.CellFormat.ReadFrom(CloneInto<W.TableCellProperties>(source.TableStyleConditionalFormattingTableCellProperties));
                yield return conditional;
            }
        }
    }

    /// <summary>
    /// The underlying OpenXML element, for anything this view does not expose.
    /// </summary>
    public W.Style ToOpenXml() =>
        element;

    static T CloneInto<T>(OpenXmlElement? source)
        where T : OpenXmlElement, new()
    {
        var target = new T();
        if (source == null)
        {
            return target;
        }

        foreach (var child in source.ChildElements)
        {
            target.AppendChild(child.CloneNode(true));
        }

        return target;
    }

    /// <summary>
    /// A readable form, for logs and debugging rather than for the file.
    /// </summary>
    public override string ToString() =>
        Id;
}
