namespace OpenXmlKit.Word;

/// <summary>
/// Something that holds block content — paragraphs and tables. The body, a header, a footer, a
/// footnote.
/// </summary>
public abstract class BlockContainer
{
    readonly OpenXmlElement container;

    // Wrappers created through this container, kept so that a flush reaches everything the caller
    // has built. The elements themselves are appended as they are added; what is deferred is only
    // the formatting, which a caller is free to set at any point before the document is saved.
    readonly List<Paragraph> paragraphs = [];
    readonly List<Table> tables = [];

    private protected BlockContainer(OpenXmlElement container, Document? document)
    {
        this.container = container;
        Document = document;
    }

    internal Document? Document { get; }

    /// <summary>
    /// Adds a paragraph, optionally holding plain text, and returns it.
    /// </summary>
    public Paragraph AddParagraph(string? text = null)
    {
        var paragraph = new Paragraph(text);
        Append(paragraph);
        return paragraph;
    }

    /// <summary>
    /// Adds an empty table and returns it.
    /// </summary>
    public Table AddTable()
    {
        var table = new Table();
        Append(table);
        return table;
    }

    /// <summary>
    /// Appends a table already built elsewhere.
    /// </summary>
    public void Append(Table table)
    {
        Insert(table.Element);
        tables.Add(table);
        InsertParagraphAfterTable();
    }

    /// <summary>
    /// Appends a paragraph built elsewhere.
    /// </summary>
    public void Append(Paragraph paragraph)
    {
        Insert(paragraph.Element);
        paragraphs.Add(paragraph);
    }

    /// <summary>
    /// Writes any formatting set on this container's content into the underlying elements.
    /// </summary>
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
    }

    /// <summary>
    /// Appends an element this library does not model.
    /// </summary>
    public void AppendElement(OpenXmlElement child) =>
        Insert(child);

    /// <summary>
    /// Where new content goes. The body overrides this to keep its section properties last.
    /// </summary>
    private protected virtual void Insert(OpenXmlElement child) =>
        container.AppendChild(child);

    // Two tables with nothing between them are read as one table with all the rows, and a container
    // that ends on a table is repaired by Word. A paragraph after each one settles both.
    void InsertParagraphAfterTable() =>
        Insert(new W.Paragraph());
}
