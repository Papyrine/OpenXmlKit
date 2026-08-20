namespace OpenXmlKit.Word;

/// <summary>
/// The content a paragraph can hold that needs more than an element: a picture, which lives in its
/// own package part; a hyperlink, which needs a relationship; a field, which is three runs and a
/// pair of markers.
/// </summary>
/// <remarks>
/// These take the owning <see cref="Document"/> explicitly rather than reaching for an ambient one,
/// because a paragraph built standalone has none — a fragment assembled in a helper and handed to a
/// document later is a shape this library supports, and it cannot register a relationship until it
/// knows where it is going.
/// </remarks>
public partial class Paragraph
{
    /// <summary>
    /// Adds a picture, sized from the image's own header unless a size is given.
    /// </summary>
    /// <param name="document">
    /// The document the image part is added to.
    /// </param>
    /// <param name="bytes">
    /// The encoded image. PNG, JPEG, GIF and BMP can be measured; anything else needs an explicit
    /// size.
    /// </param>
    /// <param name="width">
    /// Drawn width. With only one of width and height given, the other follows the aspect ratio.
    /// </param>
    /// <param name="height">
    /// Drawn height.
    /// </param>
    /// <param name="wrap">
    /// Whether the picture sits in the line of text or floats beside it.
    /// </param>
    /// <param name="description">
    /// Alternative text, which is what a screen reader announces.
    /// </param>
    public Run AddImage(
        Document document,
        byte[] bytes,
        Length? width = null,
        Length? height = null,
        ImageWrap wrap = ImageWrap.Inline,
        string? description = null)
    {
        var format = ImageFormat.Png;
        var intrinsicWidth = Length.Zero;
        var intrinsicHeight = Length.Zero;
        if (ImageInfo.TryRead(bytes, out var info))
        {
            format = info.Format;
            intrinsicWidth = info.Width;
            intrinsicHeight = info.Height;
        }
        else if (width == null || height == null)
        {
            throw new NotSupportedException(
                "Could not read the image dimensions, so a width and a height are both required. " +
                "PNG, JPEG, GIF and BMP can be measured from their own headers.");
        }

        var (drawnWidth, drawnHeight) = Resolve(width, height, intrinsicWidth, intrinsicHeight);

        var part = document.MainPart.AddImagePart(Images.PartTypeOf(format), document.NextImageRelationshipId());
        using (var stream = new MemoryStream(bytes))
        {
            part.FeedData(stream);
        }

        var run = AddRun();
        run.AppendElement(
            Images.Build(
                document.MainPart.GetIdOfPart(part),
                document.NextDrawingId(),
                description ?? "Image",
                drawnWidth,
                drawnHeight,
                wrap,
                description));
        return run;
    }

    public Run AddImage(
        Document document,
        Stream stream,
        Length? width = null,
        Length? height = null,
        ImageWrap wrap = ImageWrap.Inline,
        string? description = null)
    {
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return AddImage(document, buffer.ToArray(), width, height, wrap, description);
    }

    // Giving one dimension and letting the other follow is what most callers want, and getting it
    // wrong is a stretched picture rather than an error, so it is worth doing here.
    static (Length Width, Length Height) Resolve(Length? width, Length? height, Length intrinsicWidth, Length intrinsicHeight)
    {
        if (width is { } bothWidth &&
            height is { } bothHeight)
        {
            return (bothWidth, bothHeight);
        }

        var ratio = intrinsicWidth.TotalPoints <= 0
            ? 1
            : intrinsicHeight.TotalPoints / intrinsicWidth.TotalPoints;

        if (width is { } onlyWidth)
        {
            return (onlyWidth, Length.FromPoints(onlyWidth.TotalPoints * ratio));
        }

        if (height is { } onlyHeight)
        {
            var scale = ratio <= 0 ? 1 : 1 / ratio;
            return (Length.FromPoints(onlyHeight.TotalPoints * scale), onlyHeight);
        }

        return (intrinsicWidth, intrinsicHeight);
    }

    /// <summary>
    /// Adds a link to a web address.
    /// </summary>
    /// <remarks>
    /// The link text takes the Hyperlink character style, whose definition is written into the
    /// document if it is not there already — otherwise the link renders as ordinary text and only
    /// reveals itself on hover.
    /// </remarks>
    public Paragraph AddLink(Document document, string url, string text, string? styleId = "Hyperlink")
    {
        var relationship = document.MainPart.AddHyperlinkRelationship(new(url, UriKind.RelativeOrAbsolute), true);
        var link = new W.Hyperlink
        {
            Id = relationship.Id
        };
        return AppendLink(document, link, text, styleId);
    }

    /// <summary>
    /// Adds a link to a bookmark in the same document.
    /// </summary>
    public Paragraph AddAnchorLink(Document document, string bookmarkName, string text, string? styleId = "Hyperlink")
    {
        var link = new W.Hyperlink
        {
            Anchor = bookmarkName
        };
        return AppendLink(document, link, text, styleId);
    }

    Paragraph AppendLink(Document document, W.Hyperlink link, string text, string? styleId)
    {
        var run = new Run();
        run.Append(text);
        if (styleId != null)
        {
            if (styleId == "Hyperlink")
            {
                document.Styles.EnsureBuiltIn(BuiltInStyle.Hyperlink);
            }

            run.Font.StyleId = styleId;
        }

        run.Flush();
        link.AppendChild(run.Element);
        AppendElement(link);
        return this;
    }

    /// <summary>
    /// Wraps content in a bookmark, so a link or a cross-reference can point at it.
    /// </summary>
    /// <remarks>
    /// A bookmark name has a character set and a length limit a title would not respect, so a name
    /// derived from user text needs sanitising first — <see cref="Bookmarks.Sanitise"/> does that.
    /// </remarks>
    public Paragraph AddBookmark(Document document, string name, Action<Paragraph>? content = null)
    {
        var id = document.NextBookmarkId();
        AppendElement(
            new W.BookmarkStart
            {
                Id = id.ToString(CultureInfo.InvariantCulture),
                Name = name
            });
        content?.Invoke(this);
        AppendElement(
            new W.BookmarkEnd
            {
                Id = id.ToString(CultureInfo.InvariantCulture)
            });
        return this;
    }

    /// <summary>
    /// Adds a field — a value Word computes, such as a page number or a cross-reference.
    /// </summary>
    /// <param name="code">
    /// The field instruction, as it appears between the field markers. <c>PAGE</c>,
    /// <c>PAGEREF du3 \h</c>, <c>TOC \o "1-3" \h</c>.
    /// </param>
    /// <param name="cachedValue">
    /// What to show until Word recalculates. Worth supplying: without a cached value Word shows a
    /// placeholder and asks the reader for permission to update fields on open, so a document meant
    /// to arrive finished should carry the computed text.
    /// </param>
    public Paragraph AddField(string code, string? cachedValue = null)
    {
        AppendElement(
            new W.Run(
                new W.FieldChar
                {
                    FieldCharType = W.FieldCharValues.Begin
                }));
        AppendElement(
            new W.Run(
                new W.FieldCode(code)
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
        AppendElement(
            new W.Run(
                new W.FieldChar
                {
                    FieldCharType = W.FieldCharValues.Separate
                }));
        AppendElement(
            new W.Run(
                new W.Text(cachedValue ?? "")
                {
                    Space = SpaceProcessingModeValues.Preserve
                }));
        AppendElement(
            new W.Run(
                new W.FieldChar
                {
                    FieldCharType = W.FieldCharValues.End
                }));
        return this;
    }

    /// <summary>
    /// Adds the current page number.
    /// </summary>
    public Paragraph AddPageNumber() =>
        AddField("PAGE");

    /// <summary>
    /// Adds a reference to the page a bookmark is on.
    /// </summary>
    public Paragraph AddPageReference(string bookmarkName, string? cachedValue = null) =>
        AddField($"PAGEREF {bookmarkName} \\h", cachedValue);

    /// <summary>
    /// Adds a footnote, and the reference mark that points at it.
    /// </summary>
    public Paragraph AddFootnote(Document document, Action<BlockContainer> content)
    {
        var id = document.AddFootnote(content);
        var reference = new Run
        {
            Font =
            {
                StyleId = "FootnoteReference"
            }
        };
        document.Styles.EnsureBuiltIn(BuiltInStyle.FootnoteReference);
        reference.Flush();
        reference.Element.AppendChild(
            new W.FootnoteReference
            {
                Id = id
            });
        AppendElement(reference.Element);
        return this;
    }
}
