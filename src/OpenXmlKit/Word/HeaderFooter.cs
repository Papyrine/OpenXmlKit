namespace OpenXmlKit.Word;

/// <summary>
/// The content of one header or footer.
/// </summary>
public class HeaderFooter :
    BlockContainer
{
    readonly OpenXmlElement element;

    internal HeaderFooter(W.Header header, Document document)
        : base(header, document) =>
        element = header;

    internal HeaderFooter(W.Footer footer, Document document)
        : base(footer, document) =>
        element = footer;

    /// <summary>
    /// Adds a paragraph of plain text, optionally with a paragraph style.
    /// </summary>
    /// <remarks>
    /// Word's own Header and Footer styles carry the tab stops that centre and right-align the
    /// three usual slots, so naming one is generally better than setting alignment directly.
    /// </remarks>
    public HeaderFooter Paragraph(string? text, string? styleId = null)
    {
        var paragraph = AddParagraph(text);
        if (styleId != null)
        {
            paragraph.Style(styleId);
        }

        return this;
    }

    public HeaderFooter AddParagraph(Action<Paragraph> configure)
    {
        configure(AddParagraph());
        return this;
    }

    public HeaderFooter AddTable(Action<Table> configure)
    {
        configure(AddTable());
        return this;
    }

    /// <summary>
    /// The underlying OpenXML element.
    /// </summary>
    public OpenXmlElement ToOpenXml() =>
        element;
}
