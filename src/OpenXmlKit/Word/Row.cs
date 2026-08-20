namespace OpenXmlKit.Word;

/// <summary>
/// A table row, and the cells it contains.
/// </summary>
public class Row
{
    readonly W.TableRow element;
    RowFormat? format;
    readonly List<Cell> cells = [];

    public Row() =>
        element = new();

    internal Row(W.TableRow element)
    {
        this.element = element;
        if (element.TableRowProperties is { } properties)
        {
            format = new();
            format.ReadFrom(properties);
        }
    }

    public RowFormat Format => format ??= new();

    public IEnumerable<Cell> Cells
    {
        get
        {
            Flush();
            foreach (var cell in element.Elements<W.TableCell>())
            {
                yield return new(cell);
            }
        }
    }

    public Cell AddCell(string? text = null)
    {
        var cell = new Cell();
        if (text != null)
        {
            cell.AddParagraph(text);
        }

        element.AppendChild(cell.Element);
        cells.Add(cell);
        return cell;
    }

    public Row AddCell(Action<Cell> configure)
    {
        configure(AddCell());
        return this;
    }

    /// <summary>
    /// Adds a cell of plain text.
    /// </summary>
    public Row Cell(string? text)
    {
        AddCell(text);
        return this;
    }

    /// <summary>
    /// Adds a cell of a given width, configured by the caller.
    /// </summary>
    public Row Cell(Width width, Action<Cell> configure)
    {
        var cell = AddCell();
        cell.Format.Width = width;
        configure(cell);
        return this;
    }

    /// <summary>
    /// Adds a cell of a given width holding one paragraph of plain text.
    /// </summary>
    public Row Cell(Width width, string? text, string? styleId = null)
    {
        var cell = AddCell();
        cell.Format.Width = width;
        cell.Paragraph(text, styleId);
        return this;
    }

    /// <summary>
    /// Repeats this row at the top of every page the table spans.
    /// </summary>
    public Row Header(Toggle value = default)
    {
        Format.IsHeader = value.IsSet ? value : Toggle.On;
        return this;
    }

    public Row Height(Length height, RowHeightRule rule = RowHeightRule.AtLeast)
    {
        Format.Height = height;
        Format.HeightRule = rule;
        return this;
    }

    /// <summary>
    /// Configures the row formatting.
    /// </summary>
    public Row Formatting(Action<RowFormat> configure)
    {
        configure(Format);
        return this;
    }

    public W.TableRow ToOpenXml()
    {
        Flush();
        return element;
    }

    internal W.TableRow Element => element;

    internal void Flush()
    {
        foreach (var cell in cells)
        {
            cell.Flush();
        }

        if (format != null)
        {
            // trPr leads the row, ahead of any cell. Assigning the typed property puts it there
            // however many cells are already in place.
            element.TableRowProperties = format.ToProperties();
        }
    }
}
