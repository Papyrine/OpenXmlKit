namespace OpenXmlKit.Word;

/// <summary>
/// Formatting for one table cell.
/// </summary>
public class CellFormat :
    ICellFormatView
{
    /// <summary>
    /// The cell's preferred width. A percentage is of the table, not the page.
    /// </summary>
    /// <remarks>
    /// Preferred rather than absolute: Word reconciles it against the table's own width and the
    /// grid, and under <see cref="TableLayout.Autofit"/> treats it as a hint. Declare
    /// <see cref="Table.ColumnWidths"/> and <see cref="TableFormat.Layout"/> together when the
    /// widths have to be honoured exactly.
    /// </remarks>
    public Width Width { get; set; } = Width.Auto;

    /// <summary>
    /// How many grid columns this cell covers. 1 is a normal cell.
    /// </summary>
    public int ColumnSpan { get; set; } = 1;

    /// <summary>
    /// This cell's part in a vertical merge. The top cell of the run is
    /// <see cref="CellMerge.Restart"/> and holds the content; the ones below it are
    /// <see cref="CellMerge.Continue"/> and must still be present, empty, for the row to have its
    /// full complement of cells.
    /// </summary>
    public CellMerge VerticalMerge { get; set; }

    /// <summary>
    /// The cell own borders, which win over the ones the table draws.
    /// </summary>
    public Borders Borders { get; } = new();

    /// <summary>
    /// A fill behind the cell.
    /// </summary>
    public Shading Shading { get; } = new();

    /// <summary>
    /// Where the content sits when the row is taller than it is.
    /// </summary>
    public VerticalAlignment? VerticalAlignment { get; set; }

    /// <summary>
    /// Which way the text runs.
    /// </summary>
    public TextDirection? TextDirection { get; set; }

    /// <summary>
    /// Padding inside the left edge, overriding the table default.
    /// </summary>
    public Length? LeftMargin { get; set; }

    /// <summary>
    /// Padding inside the right edge.
    /// </summary>
    public Length? RightMargin { get; set; }

    /// <summary>
    /// Padding inside the top edge.
    /// </summary>
    public Length? TopMargin { get; set; }

    /// <summary>
    /// Padding inside the bottom edge.
    /// </summary>
    public Length? BottomMargin { get; set; }

    /// <summary>
    /// Keeps the cell's content on one line, letting it overflow rather than wrap.
    /// </summary>
    public Toggle NoWrap { get; set; }

    /// <summary>
    /// Squeezes the content horizontally to fit the cell width.
    /// </summary>
    public Toggle FitText { get; set; }

    IBordersView ICellFormatView.Borders => Borders;
    IShadingView ICellFormatView.Shading => Shading;

    /// <summary>
    /// Sets all four margins at once.
    /// </summary>
    public void SetMargins(Length horizontal, Length vertical)
    {
        LeftMargin = horizontal;
        RightMargin = horizontal;
        TopMargin = vertical;
        BottomMargin = vertical;
    }

    /// <summary>
    /// Whether anything at all is stated. An empty format writes no properties element,
    /// leaving the style hierarchy to resolve every value.
    /// </summary>
    public bool IsEmpty =>
        Width.IsAuto &&
        ColumnSpan == 1 &&
        VerticalMerge == CellMerge.None &&
        Borders.IsEmpty &&
        Shading.IsEmpty &&
        VerticalAlignment == null &&
        TextDirection == null &&
        LeftMargin == null &&
        RightMargin == null &&
        TopMargin == null &&
        BottomMargin == null &&
        !NoWrap.IsSet &&
        !FitText.IsSet;

    /// <summary>
    /// An independent copy, so the two can diverge.
    /// </summary>
    public CellFormat Clone()
    {
        var clone = new CellFormat();
        clone.CopyFrom(this);
        return clone;
    }

    /// <summary>
    /// Overwrites every property with the other value, stated or not.
    /// </summary>
    public void CopyFrom(CellFormat other)
    {
        Width = other.Width;
        ColumnSpan = other.ColumnSpan;
        VerticalMerge = other.VerticalMerge;
        Borders.CopyFrom(other.Borders);
        Shading.CopyFrom(other.Shading);
        VerticalAlignment = other.VerticalAlignment;
        TextDirection = other.TextDirection;
        LeftMargin = other.LeftMargin;
        RightMargin = other.RightMargin;
        TopMargin = other.TopMargin;
        BottomMargin = other.BottomMargin;
        NoWrap = other.NoWrap;
        FitText = other.FitText;
    }

    /// <summary>
    /// Returns every property to unstated, so the style hierarchy resolves the lot.
    /// </summary>
    public void Clear() =>
        CopyFrom(new());

    internal W.TableCellProperties? ToProperties()
    {
        if (IsEmpty)
        {
            return null;
        }

        var properties = new W.TableCellProperties();

        if (!Width.IsAuto)
        {
            properties.TableCellWidth = WidthElement.CellWidth(Width);
        }

        if (ColumnSpan > 1)
        {
            properties.GridSpan = new()
            {
                Val = ColumnSpan
            };
        }

        if (VerticalMerge != CellMerge.None)
        {
            var merge = new W.VerticalMerge();
            // A continuation carries no val at all in what Word writes; only the restart is
            // spelled out. Emitting val="continue" is legal but differs from the reference output.
            if (VerticalMerge == CellMerge.Restart)
            {
                merge.Val = W.MergedCellValues.Restart;
            }

            properties.VerticalMerge = merge;
        }

        if (Borders.ToCellBorders() is { } borders)
        {
            properties.TableCellBorders = borders;
        }

        if (Shading.ToOpenXml() is { } shading)
        {
            properties.Shading = shading;
        }

        properties.NoWrap = Toggles.OnOffOnly<W.NoWrap>(NoWrap);

        if (LeftMargin != null ||
            RightMargin != null ||
            TopMargin != null ||
            BottomMargin != null)
        {
            var margin = new W.TableCellMargin();
            if (TopMargin is { } top)
            {
                margin.TopMargin = new()
                {
                    Width = top.Twips.ToString(CultureInfo.InvariantCulture),
                    Type = W.TableWidthUnitValues.Dxa
                };
            }

            if (LeftMargin is { } left)
            {
                margin.StartMargin = new()
                {
                    Width = left.Twips.ToString(CultureInfo.InvariantCulture),
                    Type = W.TableWidthUnitValues.Dxa
                };
            }

            if (BottomMargin is { } bottom)
            {
                margin.BottomMargin = new()
                {
                    Width = bottom.Twips.ToString(CultureInfo.InvariantCulture),
                    Type = W.TableWidthUnitValues.Dxa
                };
            }

            if (RightMargin is { } right)
            {
                margin.EndMargin = new()
                {
                    Width = right.Twips.ToString(CultureInfo.InvariantCulture),
                    Type = W.TableWidthUnitValues.Dxa
                };
            }

            properties.TableCellMargin = margin;
        }

        if (TextDirection is { } textDirection &&
            textDirection != Word.TextDirection.Horizontal)
        {
            properties.TextDirection = new()
            {
                Val = textDirection.ToOpenXml()
            };
        }

        properties.TableCellFitText = Toggles.OnOffOnly<W.TableCellFitText>(FitText);

        if (VerticalAlignment is { } verticalAlignment)
        {
            properties.TableCellVerticalAlignment = new()
            {
                Val = verticalAlignment.ToOpenXml()
            };
        }

        return properties;
    }

    internal void ReadFrom(W.TableCellProperties? properties)
    {
        if (properties == null)
        {
            return;
        }

        Width = WidthElement.Read(properties.TableCellWidth);

        if (properties.GridSpan?.Val is { HasValue: true } span)
        {
            ColumnSpan = span.Value;
        }

        if (properties.VerticalMerge is { } merge)
        {
            VerticalMerge = merge.Val is { HasValue: true } value && value.Value == W.MergedCellValues.Restart
                ? CellMerge.Restart
                : CellMerge.Continue;
        }

        Borders.ReadFrom(properties.TableCellBorders);
        Shading.ReadFrom(properties.Shading);

        if (properties.TableCellVerticalAlignment?.Val is { HasValue: true } alignment)
        {
            VerticalAlignment = Map.ToVerticalAlignment(alignment.Value);
        }

        if (properties.TextDirection?.Val is { HasValue: true } direction)
        {
            TextDirection = Map.ToTextDirection(direction.Value);
        }

        if (properties.TableCellMargin is { } margins)
        {
            TopMargin = ReadMargin(margins.TopMargin);
            LeftMargin = ReadMargin(margins.StartMargin);
            BottomMargin = ReadMargin(margins.BottomMargin);
            RightMargin = ReadMargin(margins.EndMargin);
        }

        NoWrap = Toggles.Read(properties.NoWrap);
        FitText = Toggles.Read(properties.TableCellFitText);
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
}
