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
    /// <remarks>
    /// Built from <see cref="AllRuns"/> rather than <see cref="Runs"/>, because a hyperlink holds
    /// its runs inside itself: reading only the paragraph's direct children drops the link text,
    /// which is content this library's own <see cref="Paragraph.AddLink"/> writes.
    /// </remarks>
    public string Text
    {
        get
        {
            var builder = new StringBuilder();
            foreach (var run in AllRuns)
            {
                builder.Append(run.Text);
            }

            return builder.ToString();
        }
    }

    /// <summary>
    /// The runs sitting directly in the paragraph. A run inside a hyperlink is not one of these —
    /// <see cref="AllRuns"/> is.
    /// </summary>
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
    /// <summary>
    /// The links in the paragraph.
    /// </summary>
    public IEnumerable<HyperlinkView> Hyperlinks
    {
        get
        {
            foreach (var link in element.Descendants<W.Hyperlink>())
            {
                yield return new(link);
            }
        }
    }

    /// <summary>
    /// The pictures in the paragraph.
    /// </summary>
    public IEnumerable<ImageView> Images
    {
        get
        {
            foreach (var drawing in element.Descendants<W.Drawing>())
            {
                yield return new(drawing);
            }
        }
    }

    /// <summary>
    /// The fields in the paragraph, each with its instruction and its cached result.
    /// </summary>
    /// <remarks>
    /// Both forms are read: the simple field, which is one element, and the complex field, which
    /// is a run of markers around an instruction and a result. A field nested inside another
    /// field's result is not distinguished from the outer one.
    /// </remarks>
    public IEnumerable<FieldView> Fields
    {
        get
        {
            foreach (var simple in element.Descendants<W.SimpleField>())
            {
                var text = new StringBuilder();
                foreach (var run in simple.Descendants<W.Run>())
                {
                    text.Append(new RunView(run).Text);
                }

                yield return new(simple.Instruction?.Value?.Trim() ?? "", text.ToString());
            }

            StringBuilder? code = null;
            StringBuilder? value = null;
            var inResult = false;
            foreach (var run in element.Descendants<W.Run>())
            {
                if (run.Ancestors<W.SimpleField>().Any())
                {
                    continue;
                }

                foreach (var child in run.ChildElements)
                {
                    switch (child)
                    {
                        case W.FieldChar { FieldCharType: { HasValue: true } type }:
                            if (type.Value == W.FieldCharValues.Begin)
                            {
                                code = new();
                                value = new();
                                inResult = false;
                            }
                            else if (type.Value == W.FieldCharValues.Separate)
                            {
                                inResult = true;
                            }
                            else if (type.Value == W.FieldCharValues.End)
                            {
                                if (code != null)
                                {
                                    yield return new(code.ToString().Trim(), value!.ToString());
                                }

                                code = null;
                                value = null;
                                inResult = false;
                            }

                            break;
                        case W.FieldCode instruction when code != null && !inResult:
                            code.Append(instruction.Text);
                            break;
                        case W.Text result when inResult && value != null:
                            value.Append(result.Text);
                            break;
                    }
                }
            }
        }
    }

    /// <summary>
    /// The ids of the footnotes referenced from this paragraph, for looking up in
    /// <see cref="DocumentView.Footnotes"/>.
    /// </summary>
    public IEnumerable<int> FootnoteReferences
    {
        get
        {
            foreach (var reference in element.Descendants<W.FootnoteReference>())
            {
                if (reference.Id?.Value is { } id)
                {
                    yield return (int) id;
                }
            }
        }
    }

    /// <summary>
    /// The names of the bookmarks that start in this paragraph.
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

    /// <summary>
    /// A readable form, for logs and debugging rather than for the file.
    /// </summary>
    public override string ToString() =>
        Text;
}
