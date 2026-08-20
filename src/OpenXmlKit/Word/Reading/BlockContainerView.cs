namespace OpenXmlKit.Word;

/// <summary>
/// Block content in a document being read — the body, a header, a footer, a footnote.
/// </summary>
public readonly struct BlockContainerView
{
    readonly OpenXmlElement container;

    internal BlockContainerView(OpenXmlElement container) =>
        this.container = container;

    public IEnumerable<ParagraphView> Paragraphs
    {
        get
        {
            foreach (var paragraph in container.Elements<W.Paragraph>())
            {
                yield return new(paragraph);
            }
        }
    }

    public IEnumerable<TableView> Tables
    {
        get
        {
            foreach (var table in container.Elements<W.Table>())
            {
                yield return new(table);
            }
        }
    }

    /// <summary>
    /// Every paragraph, including those inside tables.
    /// </summary>
    public IEnumerable<ParagraphView> AllParagraphs
    {
        get
        {
            foreach (var paragraph in container.Descendants<W.Paragraph>())
            {
                yield return new(paragraph);
            }
        }
    }

    /// <summary>
    /// The plain text, blocks joined by newlines.
    /// </summary>
    public string Text =>
        string.Join("\n", Paragraphs.Select(_ => _.Text));

    /// <summary>
    /// The underlying OpenXML element, for anything this view does not expose.
    /// </summary>
    public OpenXmlElement ToOpenXml() =>
        container;
}
