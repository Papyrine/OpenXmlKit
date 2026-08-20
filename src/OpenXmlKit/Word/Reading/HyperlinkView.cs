namespace OpenXmlKit.Word;

/// <summary>
/// A link in a document being read.
/// </summary>
public readonly struct HyperlinkView
{
    readonly W.Hyperlink element;

    internal HyperlinkView(W.Hyperlink element) =>
        this.element = element;

    /// <summary>
    /// The link's visible text.
    /// </summary>
    public string Text
    {
        get
        {
            var builder = new StringBuilder();
            foreach (var run in Runs)
            {
                builder.Append(run.Text);
            }

            return builder.ToString();
        }
    }

    /// <summary>
    /// The runs the link wraps.
    /// </summary>
    public IEnumerable<RunView> Runs
    {
        get
        {
            foreach (var run in element.Descendants<W.Run>())
            {
                yield return new(run);
            }
        }
    }

    /// <summary>
    /// The address the link points at, for a link out of the document.
    /// </summary>
    /// <remarks>
    /// Held as a relationship rather than on the element, so this resolves it against the part the
    /// link was read from. Null for a link to a bookmark, which uses <see cref="Anchor"/> instead.
    /// </remarks>
    public string? Url
    {
        get
        {
            if (element.Id?.Value is not { } id ||
                PartLookup.Of(element) is not { } part)
            {
                return null;
            }

            foreach (var relationship in part.HyperlinkRelationships)
            {
                if (relationship.Id == id)
                {
                    return relationship.Uri.OriginalString;
                }
            }

            return null;
        }
    }

    /// <summary>
    /// The bookmark the link points at, for a link within the document.
    /// </summary>
    public string? Anchor => element.Anchor?.Value;

    /// <summary>
    /// The text Word shows on hover, where the link states one.
    /// </summary>
    public string? Tooltip => element.Tooltip?.Value;

    /// <summary>
    /// The underlying OpenXML element, for anything this view does not expose.
    /// </summary>
    public W.Hyperlink ToOpenXml() =>
        element;

    /// <summary>
    /// A readable form, for logs and debugging rather than for the file.
    /// </summary>
    public override string ToString() =>
        Text;
}
