namespace OpenXmlKit.Word;

/// <summary>
/// Formatting for a table as a whole.
/// </summary>
public class TableFormat :
    ITableFormatView
{
    /// <summary>
    /// The id of the table style to apply.
    /// </summary>
    /// <remarks>
    /// The style has to exist in the document it lands in. <see cref="Styles.EnsureBuiltIn(BuiltInStyle)"/>
    /// supplies Word's own built-ins, which a generated document does not otherwise carry; a
    /// template's own style has no stock definition to fall back on, and naming a missing one
    /// renders as an unstyled table rather than as anything invented here.
    /// </remarks>
    public string? StyleId { get; set; }

    /// <summary>
    /// The preferred width of the table.
    /// </summary>
    public Width Width { get; set; } = Width.Auto;

    /// <summary>
    /// Where the table sits when it is narrower than the text width.
    /// </summary>
    public TableAlignment? Alignment { get; set; }

    /// <summary>
    /// How column widths are decided.
    /// </summary>
    public TableLayout? Layout { get; set; }

    /// <summary>
    /// Space between the left margin and the table.
    /// </summary>
    public Length? Indent { get; set; }

    /// <summary>
    /// The borders drawn around and between cells.
    /// </summary>
    public Borders Borders { get; } = new();

    /// <summary>
    /// A fill behind the whole table.
    /// </summary>
    public Shading Shading { get; } = new();

    /// <summary>
    /// The default padding inside every cell, which an individual cell can override.
    /// </summary>
    public Length? DefaultLeftMargin { get; set; }

    /// <summary>
    /// The default padding inside the right edge of every cell.
    /// </summary>
    public Length? DefaultRightMargin { get; set; }

    /// <summary>
    /// The default padding inside the top edge of every cell.
    /// </summary>
    public Length? DefaultTopMargin { get; set; }

    /// <summary>
    /// The default padding inside the bottom edge of every cell.
    /// </summary>
    public Length? DefaultBottomMargin { get; set; }

    /// <summary>
    /// Which conditional parts of the table style apply — the header row shading, the banding, and
    /// so on.
    /// </summary>
    public TableLook Look { get; } = new();

    /// <summary>
    /// The table's title, as assistive technology announces it.
    /// </summary>
    /// <remarks>
    /// Word surfaces this and <see cref="Description"/> as "Alt Text" on a table. A table used for
    /// layout rather than data wants neither; a data table wants both, and a document that has to
    /// meet an accessibility standard will not without them.
    /// </remarks>
    public string? Caption { get; set; }

    /// <summary>
    /// A longer description of what the table contains, for the same audience as
    /// <see cref="Caption"/>.
    /// </summary>
    public string? Description { get; set; }

    IBordersView ITableFormatView.Borders => Borders;
    IShadingView ITableFormatView.Shading => Shading;
    ITableLookView ITableFormatView.Look => Look;

    /// <summary>
    /// Sets the default cell padding on both axes at once.
    /// </summary>
    public void SetDefaultMargins(Length horizontal, Length vertical)
    {
        DefaultLeftMargin = horizontal;
        DefaultRightMargin = horizontal;
        DefaultTopMargin = vertical;
        DefaultBottomMargin = vertical;
    }

    /// <summary>
    /// Whether anything at all is stated. An empty format writes no properties element,
    /// leaving the style hierarchy to resolve every value.
    /// </summary>
    public bool IsEmpty =>
        StyleId == null &&
        Width.IsAuto &&
        Alignment == null &&
        Layout == null &&
        Indent == null &&
        Borders.IsEmpty &&
        Shading.IsEmpty &&
        DefaultLeftMargin == null &&
        DefaultRightMargin == null &&
        DefaultTopMargin == null &&
        DefaultBottomMargin == null &&
        Look.IsDefault &&
        Caption == null &&
        Description == null;

    /// <summary>
    /// An independent copy, so the two can diverge.
    /// </summary>
    public TableFormat Clone()
    {
        var clone = new TableFormat();
        clone.CopyFrom(this);
        return clone;
    }

    /// <summary>
    /// Overwrites every property with the other value, stated or not.
    /// </summary>
    public void CopyFrom(TableFormat other)
    {
        StyleId = other.StyleId;
        Width = other.Width;
        Alignment = other.Alignment;
        Layout = other.Layout;
        Indent = other.Indent;
        Borders.CopyFrom(other.Borders);
        Shading.CopyFrom(other.Shading);
        DefaultLeftMargin = other.DefaultLeftMargin;
        DefaultRightMargin = other.DefaultRightMargin;
        DefaultTopMargin = other.DefaultTopMargin;
        DefaultBottomMargin = other.DefaultBottomMargin;
        Look.CopyFrom(other.Look);
        Caption = other.Caption;
        Description = other.Description;
    }

    /// <summary>
    /// Returns every property to unstated, so the style hierarchy resolves the lot.
    /// </summary>
    public void Clear() =>
        CopyFrom(new());

    /// <summary>
    /// The same formatting as it appears inside a style definition.
    /// </summary>
    /// <remarks>
    /// A style's table properties are CT_TblPrBase, which is CT_TblPr without tblLook: which
    /// conditional parts of a style apply is a property of the table using it, not of the style
    /// itself, so stating it here makes the styles part schema-invalid.
    /// </remarks>
    internal W.StyleTableProperties ToStyleProperties()
    {
        var target = new W.StyleTableProperties();
        if (ToProperties() is not { } source)
        {
            return target;
        }

        foreach (var child in source.ChildElements.ToList())
        {
            child.Remove();
            if (child is W.TableLook)
            {
                continue;
            }

            target.AppendChild(child);
        }

        return target;
    }

    internal W.TableProperties? ToProperties()
    {
        if (IsEmpty)
        {
            return null;
        }

        var properties = new W.TableProperties();

        if (StyleId != null)
        {
            properties.TableStyle = new()
            {
                Val = StyleId
            };
        }

        if (!Width.IsAuto)
        {
            properties.TableWidth = WidthElement.TableWidth(Width);
        }

        if (Alignment is { } alignment)
        {
            properties.TableJustification = new()
            {
                Val = ToOpenXml(alignment)
            };
        }

        if (Indent is { } indent)
        {
            properties.TableIndentation = new()
            {
                Width = indent.Twips,
                Type = W.TableWidthUnitValues.Dxa
            };
        }

        if (Borders.ToTableBorders() is { } borders)
        {
            properties.TableBorders = borders;
        }

        if (Shading.ToOpenXml() is { } shading)
        {
            properties.Shading = shading;
        }

        if (Layout is { } layout)
        {
            properties.TableLayout = new()
            {
                Type = layout.ToOpenXml()
            };
        }

        if (DefaultLeftMargin != null ||
            DefaultRightMargin != null ||
            DefaultTopMargin != null ||
            DefaultBottomMargin != null)
        {
            var margin = new W.TableCellMarginDefault();
            if (DefaultTopMargin is { } top)
            {
                margin.TopMargin = new()
                {
                    Width = top.Twips.ToString(CultureInfo.InvariantCulture),
                    Type = W.TableWidthUnitValues.Dxa
                };
            }

            if (DefaultLeftMargin is { } left)
            {
                margin.StartMargin = new()
                {
                    Width = left.Twips.ToString(CultureInfo.InvariantCulture),
                    Type = W.TableWidthUnitValues.Dxa
                };
            }

            if (DefaultBottomMargin is { } bottom)
            {
                margin.BottomMargin = new()
                {
                    Width = bottom.Twips.ToString(CultureInfo.InvariantCulture),
                    Type = W.TableWidthUnitValues.Dxa
                };
            }

            if (DefaultRightMargin is { } right)
            {
                margin.EndMargin = new()
                {
                    Width = right.Twips.ToString(CultureInfo.InvariantCulture),
                    Type = W.TableWidthUnitValues.Dxa
                };
            }

            properties.TableCellMarginDefault = margin;
        }

        if (Look.ToOpenXml() is { } look)
        {
            properties.TableLook = look;
        }

        if (Caption != null)
        {
            properties.TableCaption = new()
            {
                Val = Caption
            };
        }

        if (Description != null)
        {
            properties.TableDescription = new()
            {
                Val = Description
            };
        }

        return properties;
    }

    internal void ReadFrom(W.TableProperties? properties)
    {
        if (properties == null)
        {
            return;
        }

        StyleId = properties.TableStyle?.Val?.Value;
        Width = WidthElement.Read(properties.TableWidth);

        if (properties.TableJustification?.Val is { HasValue: true } alignment)
        {
            Alignment = ToAlignment(alignment.Value);
        }

        if (properties.TableIndentation?.Width is { HasValue: true } indent)
        {
            Indent = Length.FromTwips(indent.Value);
        }

        Borders.ReadFrom(properties.TableBorders);
        Shading.ReadFrom(properties.Shading);

        if (properties.TableLayout?.Type is { HasValue: true } layout)
        {
            Layout = layout.Value == W.TableLayoutValues.Fixed ? TableLayout.Fixed : TableLayout.Autofit;
        }

        if (properties.TableCellMarginDefault is { } margins)
        {
            DefaultTopMargin = ReadMargin(margins.TopMargin);
            DefaultLeftMargin = ReadMargin(margins.StartMargin);
            DefaultBottomMargin = ReadMargin(margins.BottomMargin);
            DefaultRightMargin = ReadMargin(margins.EndMargin);
        }

        Look.ReadFrom(properties.TableLook);
        Caption = properties.TableCaption?.Val?.Value;
        Description = properties.TableDescription?.Val?.Value;
    }

    static Length? ReadMargin(W.TableWidthType? margin)
    {
        if (margin?.Width is not { HasValue: true } width ||
            !double.TryParse(width.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var twips))
        {
            return null;
        }

        return Length.FromTwips(twips);
    }

    static W.TableRowAlignmentValues ToOpenXml(TableAlignment alignment) =>
        alignment switch
        {
            TableAlignment.Center => W.TableRowAlignmentValues.Center,
            TableAlignment.Right => W.TableRowAlignmentValues.Right,
            _ => W.TableRowAlignmentValues.Left
        };

    static TableAlignment ToAlignment(W.TableRowAlignmentValues value)
    {
        if (value == W.TableRowAlignmentValues.Center)
        {
            return TableAlignment.Center;
        }

        if (value == W.TableRowAlignmentValues.Right)
        {
            return TableAlignment.Right;
        }

        return TableAlignment.Left;
    }
}
