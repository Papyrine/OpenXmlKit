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
    List<TableStyleConditional>? conditionals;

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

    /// <summary>
    /// The formatting this style applies to one part of a table — the header row, the banding, a
    /// corner cell. Table styles only.
    /// </summary>
    /// <remarks>
    /// Returns the existing block for the area if the style already has one, so calling this twice
    /// for the same area configures the same block rather than adding a second.
    /// </remarks>
    public TableStyleConditional Conditional(TableStyleArea area)
    {
        var existing = Conditionals.FirstOrDefault(_ => _.Area == area);
        if (existing != null)
        {
            return existing;
        }

        var conditional = new TableStyleConditional(area);
        conditionals!.Add(conditional);
        return conditional;
    }

    /// <summary>
    /// Configures the formatting for one part of a table.
    /// </summary>
    public Style Conditional(TableStyleArea area, Action<TableStyleConditional> configure)
    {
        configure(Conditional(area));
        return this;
    }

    /// <summary>
    /// Every conditional block the style carries.
    /// </summary>
    public IReadOnlyList<TableStyleConditional> Conditionals =>
        conditionals ??= ReadConditionals();

    // Read lazily rather than in the constructor, because the blocks are rewritten wholesale on
    // flush and only a caller that asks for them should trigger that. A style opened from a
    // template and never asked keeps whatever it arrived with.
    List<TableStyleConditional> ReadConditionals()
    {
        var result = new List<TableStyleConditional>();
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
            result.Add(conditional);
        }

        return result;
    }

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
            var properties = tableFormat.ToStyleProperties();
            // tblPr precedes tblStylePr in CT_Style's sequence, so a style that already carries
            // conditional blocks needs the table properties put in front of them rather than
            // appended after.
            if (element.GetFirstChild<W.TableStyleProperties>() is { } first)
            {
                element.InsertBefore(properties, first);
            }
            else
            {
                element.Append(properties);
            }
        }

        if (conditionals != null)
        {
            foreach (var existing in element.Elements<W.TableStyleProperties>().ToList())
            {
                existing.Remove();
            }

            foreach (var conditional in conditionals)
            {
                if (!conditional.IsEmpty)
                {
                    element.Append(conditional.ToOpenXml());
                }
            }
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

    // The style-scoped and override-scoped property elements carry the same children as the
    // content-scoped ones, so reading either means copying across into the type the format objects
    // already know how to read.
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
