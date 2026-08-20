namespace OpenXmlKit.Word;

/// <summary>
/// A table cell, and the block content it contains.
/// </summary>
public class Cell
{
    readonly W.TableCell element;
    CellFormat? format;
    readonly List<Paragraph> paragraphs = [];
    readonly List<Table> tables = [];

    public Cell() =>
        element = new();

    public Cell(string? text)
        : this() =>
        AddParagraph(text);

    public CellFormat Format => format ??= new();

    public Paragraph AddParagraph(string? text = null)
    {
        var paragraph = new Paragraph(text);
        element.AppendChild(paragraph.Element);
        paragraphs.Add(paragraph);
        return paragraph;
    }

    public Cell AddParagraph(Action<Paragraph> configure)
    {
        configure(AddParagraph());
        return this;
    }

    /// <summary>
    /// Adds a paragraph of plain text, optionally with a paragraph style.
    /// </summary>
    public Cell Paragraph(string? text, string? styleId = null)
    {
        var paragraph = AddParagraph(text);
        if (styleId != null)
        {
            paragraph.Style(styleId);
        }

        return this;
    }

    /// <summary>
    /// Nests a table inside the cell.
    /// </summary>
    /// <remarks>
    /// A cell has to end with a paragraph, so one is added after the table if nothing else follows
    /// it. Word repairs a document where a cell ends on a table, and repairing changes the layout.
    /// </remarks>
    public Table AddTable()
    {
        var table = new Table();
        element.AppendChild(table.Element);
        tables.Add(table);
        return table;
    }

    public Cell AddTable(Action<Table> configure)
    {
        configure(AddTable());
        return this;
    }

    public Cell Width(Width width)
    {
        Format.Width = width;
        return this;
    }

    public Cell ColumnSpan(int columns)
    {
        Format.ColumnSpan = columns;
        return this;
    }

    public Cell Background(Color color)
    {
        Format.Shading.BackgroundColor = color;
        return this;
    }

    public Cell VerticalAlignment(VerticalAlignment alignment)
    {
        Format.VerticalAlignment = alignment;
        return this;
    }

    /// <summary>
    /// Configures the cell formatting.
    /// </summary>
    public Cell Formatting(Action<CellFormat> configure)
    {
        configure(Format);
        return this;
    }

    /// <summary>
    /// Appends an element this library does not model.
    /// </summary>
    public Cell AppendElement(OpenXmlElement child)
    {
        element.AppendChild(child);
        return this;
    }

    public W.TableCell ToOpenXml()
    {
        Flush();
        return element;
    }

    internal W.TableCell Element => element;

    internal void Flush()
    {
        foreach (var paragraph in paragraphs)
        {
            paragraph.Flush();
        }

        foreach (var table in tables)
        {
            table.Flush();
        }

        // A cell that has said nothing still needs a paragraph in it, and one that ends on a table
        // needs one after it. Neither is valid content, and Word repairs the document rather than
        // rendering it.
        if (!element.Elements<W.Paragraph>().Any() ||
            element.LastChild is W.Table)
        {
            element.AppendChild(new W.Paragraph());
        }

        if (format != null)
        {
            element.TableCellProperties = format.ToProperties();
        }
    }

    // The element as it stands, without the trailing-paragraph repair. The cursor builder writes
    // into a cell after opening it, so repairing at that point would leave an empty paragraph
    // sitting above everything it then writes.
    internal W.TableCell Container => element;
}
