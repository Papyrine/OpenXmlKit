namespace OpenXmlKit.Word;

/// <summary>
/// A table row, and the cells it contains.
/// </summary>
public class Row
{
    readonly W.TableRow element;
    RowFormat? format;
    readonly List<Cell> cells = [];

    /// <summary>
    /// A row with no cells and no formatting.
    /// </summary>
    public Row() =>
        element = new();

    /// <summary>
    /// The row formatting, applied when the document is flushed.
    /// </summary>
    public RowFormat Format => format ??= new();

    /// <summary>
    /// Adds a cell, optionally holding one paragraph of plain text, and returns it.
    /// </summary>
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

    /// <summary>
    /// Adds a cell and configures it.
    /// </summary>
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

    /// <summary>
    /// Sets the row height and how it is interpreted.
    /// </summary>
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

    /// <summary>
    /// The underlying element, flushed. The escape hatch for anything not modelled here.
    /// </summary>
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
