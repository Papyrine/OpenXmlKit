namespace OpenXmlKit.Word;

/// <summary>
/// A table, and the rows it contains.
/// </summary>
/// <remarks>
/// Self-contained: a table can be built without a document and handed to one later, or handed
/// straight to code that wants the raw element. That is what lets a report build a fragment in a
/// helper method the way it used to build a string of html.
/// </remarks>
public class Table
{
    readonly W.Table element;
    TableFormat? format;
    List<Length>? columnWidths;
    readonly List<Row> rows = [];

    public Table() =>
        element = new();

    /// <summary>
    /// Starts a new table. Equivalent to the constructor, and reads better at the head of a chain.
    /// </summary>
    public static Table Create() =>
        new();

    public TableFormat Format => format ??= new();

    /// <summary>
    /// The explicit grid column widths.
    /// </summary>
    /// <remarks>
    /// Word resolves a table's layout against the grid, not against the cell widths, so a table
    /// whose columns must come out at declared widths needs both this and
    /// <see cref="TableLayout.Fixed"/>. Left unset, a grid of empty columns is written to match the
    /// widest row, and Word fits the columns to their content.
    /// </remarks>
    public IReadOnlyList<Length>? ColumnWidths => columnWidths;

    public Row AddRow()
    {
        var row = new Row();
        element.AppendChild(row.Element);
        rows.Add(row);
        return row;
    }

    public Table AddRow(Action<Row> configure)
    {
        configure(AddRow());
        return this;
    }

    /// <summary>
    /// Adds a row and configures it.
    /// </summary>
    public Table Row(Action<Row> configure) =>
        AddRow(configure);

    /// <summary>
    /// Adds a row of plain text cells, one per value.
    /// </summary>
    public Table Row(params string?[] cells)
    {
        var row = AddRow();
        foreach (var cell in cells)
        {
            row.AddCell(cell);
        }

        return this;
    }

    /// <summary>
    /// Adds a header row of plain text cells, repeated at the top of every page the table spans.
    /// </summary>
    public Table HeaderRow(params string?[] cells)
    {
        var row = AddRow();
        row.Header();
        foreach (var cell in cells)
        {
            row.AddCell(cell);
        }

        return this;
    }

    /// <summary>
    /// Applies a table style by id.
    /// </summary>
    public Table Style(string? styleId)
    {
        Format.StyleId = styleId;
        return this;
    }

    public Table Width(Width width)
    {
        Format.Width = width;
        return this;
    }

    /// <summary>
    /// Declares the grid column widths and switches to fixed layout, so Word honours them.
    /// </summary>
    public Table Columns(params Length[] widths)
    {
        columnWidths = [.. widths];
        Format.Layout = TableLayout.Fixed;
        return this;
    }

    /// <summary>
    /// Draws the same line on every edge, inside and out.
    /// </summary>
    public Table Borders(BorderStyle style, Length? width = null, Color? color = null)
    {
        Format.Borders.SetAll(style, width, color);
        return this;
    }

    /// <summary>
    /// Configures the table formatting.
    /// </summary>
    public Table Formatting(Action<TableFormat> configure)
    {
        configure(Format);
        return this;
    }

    public W.Table ToOpenXml()
    {
        Flush();
        return element;
    }

    internal W.Table Element => element;

    // Unlike a paragraph or a cell, a table's leading children have no typed properties to assign
    // through - the SDK models them as ordinary children - so their schema order has to be produced
    // here. tblPr then tblGrid then the rows: both are removed and prepended in reverse, so each
    // lands ahead of what went before it. Rebuilding rather than patching keeps this idempotent,
    // which matters because reading the element flushes it.
    internal void Flush()
    {
        foreach (var row in rows)
        {
            row.Flush();
        }

        element.GetFirstChild<W.TableGrid>()?.Remove();
        element.PrependChild(BuildGrid());

        element.GetFirstChild<W.TableProperties>()?.Remove();
        // tblPr is required by CT_Tbl rather than optional, so a table that states no formatting
        // still gets an empty one. Without it Word reports the document as corrupt, and the
        // validator's complaint lands on tblGrid, which is the next element rather than the missing
        // one - a good way to spend an afternoon looking at the wrong thing.
        element.PrependChild(format?.ToProperties() ?? new W.TableProperties());
    }

    W.TableGrid BuildGrid()
    {
        var grid = new W.TableGrid();
        if (columnWidths is { Count: > 0 })
        {
            foreach (var width in columnWidths)
            {
                grid.Append(
                    new W.GridColumn
                    {
                        Width = width.Twips.ToString(CultureInfo.InvariantCulture)
                    });
            }

            return grid;
        }

        // No declared widths, so the grid is as wide as the widest row and carries no measurements.
        // A grid that is narrower than a row makes Word drop the surplus cells.
        var columns = 0;
        foreach (var row in element.Elements<W.TableRow>())
        {
            var count = 0;
            foreach (var cell in row.Elements<W.TableCell>())
            {
                count += cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1;
            }

            if (count > columns)
            {
                columns = count;
            }
        }

        for (var index = 0; index < columns; index++)
        {
            grid.Append(new W.GridColumn());
        }

        return grid;
    }
}
