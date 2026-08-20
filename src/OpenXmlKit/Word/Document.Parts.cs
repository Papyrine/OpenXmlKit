namespace OpenXmlKit.Word;

/// <summary>
/// The document-scoped counters and side parts: image relationships, drawing ids, bookmark ids,
/// footnotes, and the document properties.
/// </summary>
/// <remarks>
/// Every one of these is a number that has to be unique across the whole document and has to
/// continue past whatever an opened document already contains. Keeping them here is what stops a
/// second call adding content that silently takes over the first call's ids.
/// </remarks>
public sealed partial class Document
{
    int nextImageIndex;
    uint nextDrawingId;
    int nextBookmarkId = -1;
    int nextFootnoteId = -1;

    // Pinned rather than generated, so the same content produces the same package on every machine
    // and a byte comparison between two runs stays meaningful.
    internal string NextImageRelationshipId() =>
        "rImage" + (++nextImageIndex).ToString(CultureInfo.InvariantCulture);

    /// <summary>
    /// Drawing ids have to be unique within the document, and Word treats zero as unset.
    /// </summary>
    internal uint NextDrawingId() =>
        ++nextDrawingId;

    internal int NextBookmarkId()
    {
        if (nextBookmarkId < 0)
        {
            nextBookmarkId = MainPart.Document?.Body?
                .Descendants<W.BookmarkStart>()
                .Select(_ => int.TryParse(_.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) ? id : 0)
                .DefaultIfEmpty(-1)
                .Max() + 1 ?? 0;
        }

        return nextBookmarkId++;
    }

    /// <summary>
    /// Adds a footnote and returns its id, for the reference mark to point at.
    /// </summary>
    internal int AddFootnote(Action<BlockContainer> content)
    {
        var part = MainPart.FootnotesPart ?? CreateFootnotesPart();
        var footnotes = part.Footnotes!;

        var id = NextFootnoteId(footnotes);
        var footnote = new W.Footnote
        {
            Id = id
        };
        footnotes.AppendChild(footnote);

        Styles.EnsureBuiltIn(BuiltInStyle.FootnoteText);
        var container = new FootnoteBody(footnote, this);
        var paragraph = container.AddParagraph();
        paragraph.Style("FootnoteText");
        // The mark that numbers the note itself, as distinct from the reference in the body text.
        paragraph.AppendElement(
            new W.Run(new W.FootnoteReferenceMark())
            {
                RunProperties = new()
                {
                    RunStyle = new()
                    {
                        Val = "FootnoteReference"
                    }
                }
            });
        paragraph.AddRun(" ");
        content(container);
        container.Flush();
        paragraph.Flush();
        return id;
    }

    int NextFootnoteId(W.Footnotes footnotes)
    {
        if (nextFootnoteId < 0)
        {
            nextFootnoteId = (int) footnotes
                .Elements<W.Footnote>()
                .Select(_ => _.Id?.Value ?? 0)
                .DefaultIfEmpty(0)
                .Max() + 1;
            if (nextFootnoteId < 1)
            {
                nextFootnoteId = 1;
            }
        }

        return nextFootnoteId++;
    }

    // Word expects the two special notes with ids 0 and -1 to be present; a footnotes part without
    // them opens as a document with a missing separator line.
    FootnotesPart CreateFootnotesPart()
    {
        var part = MainPart.AddNewPart<FootnotesPart>("rFootnotes");
        part.Footnotes = new(
            new W.Footnote(
                new W.Paragraph(
                    new W.Run(new W.SeparatorMark())))
            {
                Type = W.FootnoteEndnoteValues.Separator,
                Id = -1
            },
            new W.Footnote(
                new W.Paragraph(
                    new W.Run(new W.ContinuationSeparatorMark())))
            {
                Type = W.FootnoteEndnoteValues.ContinuationSeparator,
                Id = 0
            });
        return part;
    }

    sealed class FootnoteBody(W.Footnote footnote, Document document) :
        BlockContainer(footnote, document);

    /// <summary>
    /// Resolves the formatting that actually applies to a piece of content, rather than the
    /// formatting written on it.
    /// </summary>
    public FormattingResolver Formatting =>
        field ??= new(this);

    /// <summary>
    /// The document's metadata — title, author, and the rest of what a file's properties dialog
    /// shows.
    /// </summary>
    public DocumentProperties Properties =>
        field ??= new(package);
}
