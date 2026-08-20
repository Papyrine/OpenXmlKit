namespace OpenXmlKit.Word;

/// <summary>
/// A table in a document being read.
/// </summary>
public readonly struct TableView
{
    readonly W.Table element;

    internal TableView(W.Table element) =>
        this.element = element;

    /// <summary>
    /// The table's text, rows joined by newlines and cells within a row by tabs.
    /// </summary>
    public string Text =>
        string.Join("\n", Rows.Select(_ => _.Text));

    /// <summary>
    /// The rows in the table, top to bottom.
    /// </summary>
    public IEnumerable<RowView> Rows
    {
        get
        {
            foreach (var row in element.Elements<W.TableRow>())
            {
                yield return new(row);
            }
        }
    }

    /// <summary>
    /// The formatting written on the table.
    /// </summary>
    public ITableFormatView Format
    {
        get
        {
            var format = new TableFormat();
            format.ReadFrom(element.GetFirstChild<W.TableProperties>());
            return format;
        }
    }

    /// <summary>
    /// The declared grid column widths. Word resolves layout against the grid rather than against
    /// the cell widths, so this is what the columns actually measure.
    /// </summary>
    public IReadOnlyList<Length> ColumnWidths
    {
        get
        {
            if (element.GetFirstChild<W.TableGrid>() is not { } grid)
            {
                return [];
            }

            var widths = new List<Length>();
            foreach (var column in grid.Elements<W.GridColumn>())
            {
                widths.Add(
                    column.Width is { HasValue: true } width &&
                    double.TryParse(width.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var twips)
                        ? Length.FromTwips(twips)
                        : Length.Zero);
            }

            return widths;
        }
    }

    /// <summary>
    /// The style named on this table, if any.
    /// </summary>
    public string? StyleId =>
        element.GetFirstChild<W.TableProperties>()?.TableStyle?.Val?.Value;

    /// <summary>
    /// The underlying OpenXML element, for anything this view does not expose.
    /// </summary>
    public W.Table ToOpenXml() =>
        element;
}
