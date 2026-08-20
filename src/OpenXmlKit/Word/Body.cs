namespace OpenXmlKit.Word;

/// <summary>
/// The main body of a document: the paragraphs and tables that make up its content, and the
/// sections they are laid out in.
/// </summary>
public class Body :
    BlockContainer
{
    readonly W.Body element;

    internal Body(W.Body element, Document? document)
        : base(element, document) =>
        this.element = element;

    /// <summary>
    /// Adds a paragraph of plain text, optionally with a paragraph style.
    /// </summary>
    public Body Paragraph(string? text, string? styleId = null)
    {
        var paragraph = AddParagraph(text);
        if (styleId != null)
        {
            paragraph.Style(styleId);
        }

        return this;
    }

    public Body AddParagraph(Action<Paragraph> configure)
    {
        configure(AddParagraph());
        return this;
    }

    public Body AddTable(Action<Table> configure)
    {
        configure(AddTable());
        return this;
    }

    /// <summary>
    /// Adds a page break, as a paragraph carrying one.
    /// </summary>
    public Body PageBreak()
    {
        AddParagraph().AppendBreak(BreakKind.Page);
        return this;
    }

    /// <summary>
    /// The section currently being written into. A document always has at least this one.
    /// </summary>
    public Section Section =>
        SectionFor(SectionProperties());

    /// <summary>
    /// Ends the current section and starts a new one, so the rest of the document can have its own
    /// page setup — a landscape run of pages in a portrait report, typically.
    /// </summary>
    /// <remarks>
    /// A section other than the last is stored on the paragraph mark that ends it rather than at
    /// the end of the body, so this moves the current properties into a paragraph of their own and
    /// leaves a copy behind for what follows. A copy rather than a fresh set, because a caller
    /// changing the orientation rarely means to reset the margins too.
    /// </remarks>
    public Section AddSection(SectionStart start = SectionStart.NewPage)
    {
        var current = SectionProperties();
        Document?.Flush();

        var breakParagraph = new W.Paragraph();
        var properties = new W.ParagraphProperties();
        properties.AppendChild(current.CloneNode(true));
        breakParagraph.ParagraphProperties = properties;
        Insert(breakParagraph);

        var carried = (W.SectionProperties) current.CloneNode(true);
        current.Remove();
        element.AppendChild(carried);

        var section = SectionFor(carried);
        section.Start = start;
        return section;
    }

    /// <summary>
    /// The underlying OpenXML element.
    /// </summary>
    public W.Body ToOpenXml()
    {
        if (Document == null)
        {
            Flush();
        }
        else
        {
            Document.Flush();
        }

        return element;
    }

    internal W.Body Element => element;

    // sectPr closes the body, so everything else goes in ahead of it. Appending past it produces a
    // document Word repairs, and repairing drops the section setup.
    private protected override void Insert(OpenXmlElement child)
    {
        if (element.GetFirstChild<W.SectionProperties>() is { } properties)
        {
            element.InsertBefore(child, properties);
            return;
        }

        element.AppendChild(child);
    }

    Section SectionFor(W.SectionProperties properties) =>
        Document?.SectionFor(properties) ?? new Section(properties, null);

    W.SectionProperties SectionProperties()
    {
        if (element.GetFirstChild<W.SectionProperties>() is { } existing)
        {
            return existing;
        }

        var properties = new W.SectionProperties();
        element.AppendChild(properties);
        return properties;
    }
}
