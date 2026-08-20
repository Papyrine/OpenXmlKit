namespace OpenXmlKit.Word;

/// <summary>
/// A paragraph in a document being read.
/// </summary>
/// <remarks>
/// A view over the element, not a copy of it and not something that can be written through. The
/// builder's <see cref="Paragraph"/> is a different type on purpose: reaching into a document you
/// opened for reading and calling <c>AddBookmark</c> on what you found should not compile, and
/// giving the two jobs one type is what would make it.
/// </remarks>
public readonly struct ParagraphView
{
    readonly W.Paragraph element;

    internal ParagraphView(W.Paragraph element) =>
        this.element = element;

    /// <summary>
    /// The paragraph's text, runs concatenated, with breaks and tabs as the characters they stand
    /// for.
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

    public IEnumerable<RunView> Runs
    {
        get
        {
            foreach (var run in element.Elements<W.Run>())
            {
                yield return new(run);
            }
        }
    }

    /// <summary>
    /// The runs inside hyperlinks as well as those directly in the paragraph.
    /// </summary>
    public IEnumerable<RunView> AllRuns
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
    /// The formatting written on this paragraph. What actually applies, once styles are resolved,
    /// is <see cref="FormattingResolver.FormatFor"/>.
    /// </summary>
    public IParagraphFormatView Format
    {
        get
        {
            var format = new ParagraphFormat();
            if (element.ParagraphProperties is { } properties)
            {
                format.ReadFrom(properties);
            }

            return format;
        }
    }

    /// <summary>
    /// The paragraph style named on this paragraph, if any.
    /// </summary>
    public string? StyleId =>
        element.ParagraphProperties?.ParagraphStyleId?.Val?.Value;

    /// <summary>
    /// The list this paragraph belongs to, if any.
    /// </summary>
    public ListMembership? List
    {
        get
        {
            if (element.ParagraphProperties?.NumberingProperties is not { } numbering ||
                numbering.NumberingId?.Val is not { HasValue: true } id)
            {
                return null;
            }

            return new(id.Value, numbering.NumberingLevelReference?.Val?.Value ?? 0);
        }
    }

    /// <summary>
    /// The names of the bookmarks starting in this paragraph.
    /// </summary>
    public IEnumerable<string> BookmarkNames
    {
        get
        {
            foreach (var bookmark in element.Descendants<W.BookmarkStart>())
            {
                if (bookmark.Name?.Value is { } name)
                {
                    yield return name;
                }
            }
        }
    }

    /// <summary>
    /// Whether this paragraph carries a section break — the section it ends is stored on its mark
    /// rather than at the end of the body.
    /// </summary>
    public bool EndsSection =>
        element.ParagraphProperties?.GetFirstChild<W.SectionProperties>() != null;

    /// <summary>
    /// The underlying OpenXML element, for anything this view does not expose.
    /// </summary>
    public W.Paragraph ToOpenXml() =>
        element;

    internal W.Paragraph Element => element;

    public override string ToString() =>
        Text;
}
