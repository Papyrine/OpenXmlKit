namespace OpenXmlKit.Word;

/// <summary>
/// A Word document. The entry point for both building one and reading one.
/// </summary>
/// <remarks>
/// Wraps the SDK rather than replacing it: <see cref="MainPart"/> and the <c>ToOpenXml</c> methods
/// on every wrapper reach the underlying elements, so anything this library does not model stays
/// available and a partial migration onto it stays possible.
/// </remarks>
public sealed partial class Document :
    IDisposable
{
    readonly WordprocessingDocument package;
    readonly MemoryStream? ownedStream;
    readonly Dictionary<W.SectionProperties, Section> sections = [];
    Body? body;
    Styles? styles;
    Numbering? numbering;
    DocumentBuilder? builder;
    bool disposed;

    Document(WordprocessingDocument package, MemoryStream? ownedStream)
    {
        this.package = package;
        this.ownedStream = ownedStream;
    }

    internal WordprocessingDocument Package => package;

    /// <summary>
    /// Starts an empty document in memory. <see cref="Save(Stream)"/> or <see cref="ToArray"/>
    /// gets the bytes out.
    /// </summary>
    public static Document Create()
    {
        var stream = new MemoryStream();
        var package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        Initialise(package);
        return new(package, stream);
    }

    /// <summary>
    /// Starts an empty document written into the given stream. The stream is written on
    /// <see cref="Dispose"/>, so the document must be disposed before the stream is read.
    /// </summary>
    public static Document Create(Stream stream)
    {
        var package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        Initialise(package);
        return new(package, null);
    }

    public static Document Create(string path)
    {
        var package = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        Initialise(package);
        return new(package, null);
    }

    /// <summary>
    /// Opens an existing document to add content to it — a branded template, typically, whose
    /// styles, headers and page setup the new content should inherit.
    /// </summary>
    /// <remarks>
    /// Building into an existing document rather than editing it: content added through
    /// <see cref="Body"/> and <see cref="Builder"/> is written, and everything already in the file
    /// is left exactly as it was.
    /// <para>
    /// Adding, not editing. There is no way to reach the content already in the file through this
    /// API — for that, open a <see cref="DocumentView"/> — so a change to it cannot be written and
    /// cannot be silently dropped either.
    /// </para>
    /// </remarks>
    /// <param name="stream">
    /// The document to add to. It must be writable and expandable: a MemoryStream built over a
    /// byte array is fixed-size and fails on the first write.
    /// </param>
    public static Document OpenForAppend(Stream stream) =>
        new(WordprocessingDocument.Open(stream, true), null);

    /// <inheritdoc cref="OpenForAppend(Stream)"/>
    public static Document OpenForAppend(string path) =>
        new(WordprocessingDocument.Open(path, true), null);

    static void Initialise(WordprocessingDocument package)
    {
        var main = package.AddMainDocumentPart();
        main.Document = new(new W.Body());
    }

    /// <summary>
    /// The underlying main document part. The escape hatch for anything not modelled here.
    /// </summary>
    public MainDocumentPart MainPart =>
        package.MainDocumentPart ??
        throw new InvalidOperationException("The document has no main part, so it is not a readable Word document.");

    /// <summary>
    /// The root document element of the main part.
    /// </summary>
    public W.Document Root =>
        MainPart.Document ??= new(new W.Body());

    public Body Body =>
        body ??= new(Root.Body ??= new(), this);

    public Styles Styles =>
        styles ??= new(this);

    public Numbering Numbering =>
        numbering ??= new(this);

    /// <summary>
    /// A cursor over the document, for building it a statement at a time rather than as a tree.
    /// </summary>
    public DocumentBuilder Builder =>
        builder ??= new(this);

    /// <summary>
    /// Writes any pending changes into the underlying elements.
    /// </summary>
    /// <remarks>
    /// Called automatically before saving and before reading an element back, so it is rarely
    /// needed directly.
    /// </remarks>
    public void Flush()
    {
        // The builder first: it may still have a paragraph or a table open, and closing those adds
        // content the body then has to flush.
        builder?.Flush();
        body?.Flush();

        foreach (var section in sections.Values)
        {
            section.Flush();
        }
    }

    public void Save()
    {
        Flush();
        Root.Save();
        styles?.Save();
        numbering?.Save();
    }

    /// <summary>
    /// Saves and copies the document into the given stream.
    /// </summary>
    public void Save(Stream stream)
    {
        Save();
        CopyTo(stream);
    }

    /// <summary>
    /// Saves and returns the document bytes.
    /// </summary>
    public byte[] ToArray()
    {
        Save();
        using var buffer = new MemoryStream();
        CopyTo(buffer);
        return buffer.ToArray();
    }

    // Cloning writes the package out without closing it, so the document stays usable afterwards
    // and this can be called more than once. What cloning does not carry is the package's core
    // properties — title, author, dates — which live outside the part graph it copies, so they are
    // put back by hand. Losing them is silent: the document opens fine, with a blank properties
    // dialog.
    void CopyTo(Stream stream)
    {
        using var clone = package.Clone(stream);
        var source = package.PackageProperties;
        var target = clone.PackageProperties;
        target.Title = source.Title;
        target.Subject = source.Subject;
        target.Creator = source.Creator;
        target.Description = source.Description;
        target.Keywords = source.Keywords;
        target.Category = source.Category;
        target.ContentStatus = source.ContentStatus;
        target.Identifier = source.Identifier;
        target.Language = source.Language;
        target.Revision = source.Revision;
        target.Version = source.Version;
        target.LastModifiedBy = source.LastModifiedBy;
        target.Created = source.Created;
        target.Modified = source.Modified;
        target.LastPrinted = source.LastPrinted;
    }

    /// <summary>
    /// Word stores the odd/even header switch once for the whole document rather than per section.
    /// </summary>
    internal void EnableEvenAndOddHeaders()
    {
        var part = MainPart.DocumentSettingsPart ?? MainPart.AddNewPart<DocumentSettingsPart>();
        part.Settings ??= new();
        if (part.Settings.GetFirstChild<W.EvenAndOddHeaders>() == null)
        {
            part.Settings.AppendChild(new W.EvenAndOddHeaders());
        }
    }

    // Section wrappers are cached per element so that page setup written through one is still
    // pending on the same instance when the document is flushed. A fresh wrapper each time would
    // quietly drop it.
    internal Section SectionFor(W.SectionProperties properties)
    {
        if (sections.TryGetValue(properties, out var existing))
        {
            return existing;
        }

        var section = new Section(properties, this);
        sections[properties] = section;
        return section;
    }

    public void Dispose()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;
        Flush();
        styles?.Save();
        numbering?.Save();
        package.Dispose();
        ownedStream?.Dispose();
    }
}
