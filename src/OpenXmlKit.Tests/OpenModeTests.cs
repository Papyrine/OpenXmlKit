/// <summary>
/// The line between the two APIs, pinned.
/// </summary>
/// <remarks>
/// Building and reading are separate types on purpose. Most of what this fixture asserts is
/// enforced by the compiler rather than at run time — <c>DocumentView.Open(...).Body.Paragraphs</c>
/// yields a <see cref="ParagraphView"/>, which has no <c>AddBookmark</c>, no <c>AddRun</c> and no
/// settable format — so what is left to test is that each API does its own job.
/// </remarks>
[TestFixture]
public class OpenModeTests
{
    [Test]
    public void ReadingReads()
    {
        using var read = DocumentView.Open(Source());

        Assert.That(read.Body.Paragraphs.Select(_ => _.Text), Is.EqualTo(["original"]));
        Assert.That(read.Text, Is.EqualTo("original"));
    }

    [Test]
    public void OpenForAppendWritesAddedContentAndLeavesTheRestAlone()
    {
        var stream = Expandable(Source());
        byte[] result;
        using (var document = Document.OpenForAppend(stream))
        {
            document.Body.Paragraph("appended");
            result = document.ToArray();
        }

        using var read = DocumentView.Open(result);
        Assert.That(read.Body.Paragraphs.Select(_ => _.Text), Is.EqualTo(["original", "appended"]));
    }

    [Test]
    public void FormattingSetOnContentAddedToAnExistingDocumentIsWritten()
    {
        var stream = Expandable(Source());
        byte[] result;
        using (var document = Document.OpenForAppend(stream))
        {
            var paragraph = document.Body.AddParagraph("added");
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            result = document.ToArray();
        }

        using var read = DocumentView.Open(result);
        var added = read.Body.Paragraphs.Last();
        Assert.That(added.Text, Is.EqualTo("added"));
        Assert.That(added.Format.Alignment, Is.EqualTo(ParagraphAlignment.Center));
    }

    [Test]
    public void AViewCanBeTakenOfADocumentStillBeingBuilt()
    {
        using var document = Document.Create();
        document.Body.Paragraph("first");

        using (var view = DocumentView.Of(document))
        {
            Assert.That(view.Body.Text, Is.EqualTo("first"));
        }

        // Disposing the view does not close the document it was taken of, so building continues.
        document.Body.Paragraph("second");
        using var after = DocumentView.Of(document);
        Assert.That(after.Body.Paragraphs.Select(_ => _.Text), Is.EqualTo(["first", "second"]));
    }

    [Test]
    public void BuildingAndReadingDoNotShareTypes()
    {
        // The guarantee, stated as a test because the compiler enforces it silently. A view is not
        // a builder and cannot be handed to one, which is what makes "change this and save it"
        // unwritable rather than merely undocumented.
        Assert.That(typeof(ParagraphView).IsAssignableTo(typeof(Paragraph)), Is.False);
        Assert.That(typeof(Paragraph).IsAssignableTo(typeof(ParagraphView)), Is.False);

        Assert.That(typeof(ParagraphView).GetMethod("AddRun"), Is.Null);
        Assert.That(typeof(ParagraphView).GetMethod("AddBookmark"), Is.Null);
        Assert.That(typeof(ParagraphView).GetMethod("AddImage"), Is.Null);

        // And the other direction: a builder has nothing to enumerate, so there is no way to reach
        // content already in the file through it.
        Assert.That(typeof(Body).GetProperty("Paragraphs"), Is.Null);
        Assert.That(typeof(Table).GetProperty("Rows"), Is.Null);
        Assert.That(typeof(Row).GetProperty("Cells"), Is.Null);
    }

    [Test]
    public void ReadFormattingIsNotSettable()
    {
        // IFontView and friends carry the same properties as their mutable counterparts with the
        // setters removed, so formatting read off a document cannot be assigned to by mistake.
        foreach (var property in typeof(IFontView).GetProperties())
        {
            Assert.That(property.CanWrite, Is.False, property.Name);
        }

        foreach (var property in typeof(IParagraphFormatView).GetProperties())
        {
            Assert.That(property.CanWrite, Is.False, property.Name);
        }
    }

    static MemoryStream Expandable(byte[] bytes)
    {
        var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        return stream;
    }

    static byte[] Source()
    {
        using var document = Document.Create();
        document.Body.Paragraph("original");
        return document.ToArray();
    }
}
