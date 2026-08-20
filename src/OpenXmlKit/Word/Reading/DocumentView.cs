namespace OpenXmlKit.Word;

/// <summary>
/// A Word document opened for reading.
/// </summary>
/// <remarks>
/// The read half of the library, and a separate API from the build half rather than the same one
/// with the setters hidden. Nothing reached through here can be written to, so the question of
/// whether a change made to something you read will reach the file does not arise — it is not
/// expressible.
/// <para>
/// Building is <see cref="Document"/>. The two do not share types: a
/// <see cref="ParagraphView"/> has no <c>AddRun</c>, and a <see cref="Paragraph"/> has no
/// enumeration to reach into.
/// </para>
/// </remarks>
public sealed class DocumentView :
    IDisposable
{
    readonly WordprocessingDocument package;
    readonly bool ownsPackage;
    bool disposed;

    DocumentView(WordprocessingDocument package, bool ownsPackage)
    {
        this.package = package;
        this.ownsPackage = ownsPackage;
    }

    public static DocumentView Open(Stream stream) =>
        new(WordprocessingDocument.Open(stream, false), true);

    public static DocumentView Open(string path) =>
        new(WordprocessingDocument.Open(path, false), true);

    /// <summary>
    /// Reads a document from its bytes.
    /// </summary>
    public static DocumentView Open(byte[] bytes) =>
        Open(new MemoryStream(bytes, false));

    /// <summary>
    /// Reads a document being built, as it currently stands.
    /// </summary>
    /// <remarks>
    /// Content still pending in the builder is flushed first, so this sees what would be saved.
    /// </remarks>
    public static DocumentView Of(Document document)
    {
        document.Flush();
        return new(document.Package, false);
    }

    /// <summary>
    /// The underlying main document part, for anything these views do not expose.
    /// </summary>
    public MainDocumentPart MainPart =>
        package.MainDocumentPart ??
        throw new InvalidOperationException("The document has no main part, so it is not a readable Word document.");

    public BlockContainerView Body =>
        new(MainPart.Document?.Body ?? new W.Body());

    /// <summary>
    /// The document's sections, in document order.
    /// </summary>
    /// <remarks>
    /// Every section but the last is stored on the paragraph mark that ends it; the last one closes
    /// the body. Walking in document order therefore means the paragraphs first.
    /// </remarks>
    public IEnumerable<SectionView> Sections
    {
        get
        {
            if (MainPart.Document?.Body is not { } body)
            {
                yield break;
            }

            foreach (var paragraph in body.Elements<W.Paragraph>())
            {
                if (paragraph.ParagraphProperties?.GetFirstChild<W.SectionProperties>() is { } properties)
                {
                    yield return new(properties, MainPart);
                }
            }

            if (body.GetFirstChild<W.SectionProperties>() is { } last)
            {
                yield return new(last, MainPart);
            }
        }
    }

    public StylesView Styles =>
        field ??= new(MainPart);

    /// <summary>
    /// Resolves the formatting that actually applies to a piece of content, rather than the
    /// formatting written on it.
    /// </summary>
    public FormattingResolver Formatting =>
        field ??= new(Styles, MainPart);

    /// <summary>
    /// The document's metadata.
    /// </summary>
    public DocumentPropertiesView Properties =>
        new(package);

    /// <summary>
    /// The document's plain text.
    /// </summary>
    public string Text =>
        Body.Text;

    public void Dispose()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;
        // A view taken of a document still being built does not own the package, and closing it
        // would take the builder down with it.
        if (ownsPackage)
        {
            package.Dispose();
        }
    }
}
