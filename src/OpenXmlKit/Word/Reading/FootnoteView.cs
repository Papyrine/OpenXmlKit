namespace OpenXmlKit.Word;

/// <summary>
/// A footnote in a document being read.
/// </summary>
public readonly struct FootnoteView
{
    readonly W.Footnote element;

    internal FootnoteView(W.Footnote element) =>
        this.element = element;

    /// <summary>
    /// The id the reference mark in the body text points at.
    /// </summary>
    public int Id => (int) (element.Id?.Value ?? 0);

    /// <summary>
    /// The note's own content.
    /// </summary>
    public BlockContainerView Body =>
        new(element);

    /// <summary>
    /// The note's text. The reference mark that numbers it is not text and does not appear here.
    /// </summary>
    public string Text => Body.Text;

    /// <summary>
    /// The underlying OpenXML element, for anything this view does not expose.
    /// </summary>
    public W.Footnote ToOpenXml() =>
        element;

    /// <summary>
    /// A readable form, for logs and debugging rather than for the file.
    /// </summary>
    public override string ToString() =>
        Text;
}
