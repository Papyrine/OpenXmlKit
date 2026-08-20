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
    /// Whether the style shows in Word's gallery.
    /// </summary>
    public bool IsQuickStyle => element.PrimaryStyle != null;

    public bool SemiHidden => element.SemiHidden != null;

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

    public override string ToString() =>
        Id;
}
