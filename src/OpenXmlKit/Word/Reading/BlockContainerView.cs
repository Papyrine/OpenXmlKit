namespace OpenXmlKit.Word;

/// <summary>
/// Block content in a document being read — the body, a header, a footer, a footnote.
/// </summary>
public readonly struct BlockContainerView
{
    readonly OpenXmlElement container;

    internal BlockContainerView(OpenXmlElement container) =>
        this.container = container;

    /// <summary>
    /// The paragraphs sitting directly in the container. A paragraph inside a table is not one.
    /// </summary>
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

    /// <summary>
    /// The tables sitting directly in the container.
    /// </summary>
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
    /// The plain text, blocks joined by newlines, in document order.
    /// </summary>
    /// <remarks>
    /// Tables are included, because a container whose whole content is a table is an ordinary
    /// shape and reporting it as empty text would be wrong. A table renders as its rows, cells
    /// separated by tabs, which is what pasting one into a text editor produces.
    /// </remarks>
    public string Text
    {
        get
        {
            var blocks = new List<string>();
            foreach (var child in container.ChildElements)
            {
                switch (child)
                {
                    case W.Paragraph paragraph:
                        blocks.Add(new ParagraphView(paragraph).Text);
                        break;
                    case W.Table table:
                        blocks.Add(new TableView(table).Text);
                        break;
                }
            }

            return string.Join("\n", blocks);
        }
    }

    /// <summary>
    /// The underlying OpenXML element, for anything this view does not expose.
    /// </summary>
    public OpenXmlElement ToOpenXml() =>
        container;
}
