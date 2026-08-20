/// <summary>
/// What <c>Text</c> means, pinned.
/// </summary>
/// <remarks>
/// Text extraction is the one read-side operation that has to see everything, because there is
/// nothing else for a caller to fall back on. Every case here is content this library's own build
/// API writes and an earlier reader dropped: a hyperlink holds its runs inside itself rather than
/// in the paragraph, and a table is not a paragraph at all.
/// </remarks>
[TestFixture]
public class TextTests
{
    [Test]
    public void HyperlinkTextIsPartOfTheParagraph()
    {
        using var document = Document.Create();
        document.Body.AddParagraph("before ")
            .AddLink(document, "https://example.org", "the link")
            .Append(" after");

        using var read = DocumentAssert.Read(document);
        var paragraph = read.Body.Paragraphs.Single();

        Assert.That(paragraph.Text, Is.EqualTo("before the link after"));
        Assert.That(read.Text, Is.EqualTo("before the link after"));

        // Runs is still the paragraph's own runs, and AllRuns is still everything. The fix is that
        // Text uses the second, not that the first changed meaning.
        Assert.That(paragraph.Runs.Count(), Is.EqualTo(2));
        Assert.That(paragraph.AllRuns.Count(), Is.EqualTo(3));
    }

    [Test]
    public void AnchorLinkTextIsPartOfTheParagraph()
    {
        using var document = Document.Create();
        document.Body.AddParagraph()
            .AddAnchorLink(document, "Section1", "see above");

        using var read = DocumentAssert.Read(document);
        Assert.That(read.Body.Paragraphs.Single().Text, Is.EqualTo("see above"));
    }

    [Test]
    public void TableTextIsPartOfTheDocument()
    {
        using var document = Document.Create();
        document.Body.AddTable(
            _ => _
                .Row("a", "b")
                .Row("c", "d"));

        using var read = DocumentAssert.Read(document);

        Assert.That(read.Body.Tables.Single().Text, Is.EqualTo("a\tb\nc\td"));

        // The trailing newline is the empty paragraph Word requires after a table — a document
        // must not end on one. It is real content that renders as a blank line, so it is reported
        // rather than trimmed: an empty paragraph a caller put there deliberately looks the same.
        Assert.That(read.Text, Is.EqualTo("a\tb\nc\td\n"));
    }

    [Test]
    public void BlocksComeOutInDocumentOrder()
    {
        using var document = Document.Create();
        document.Body.Paragraph("first");
        document.Body.AddTable(_ => _.Row("cell"));
        document.Body.Paragraph("last");

        using var read = DocumentAssert.Read(document);
        Assert.That(read.Text, Is.EqualTo("first\ncell\n\nlast"));
    }

    [Test]
    public void NestedTableTextIsPartOfTheCell()
    {
        using var document = Document.Create();
        document.Body.AddTable(
            _ => _.AddRow(
                row => row.AddCell(
                    cell =>
                    {
                        cell.AddParagraph("outer");
                        cell.AddTable(inner => inner.Row("inner"));
                    })));

        using var read = DocumentAssert.Read(document);
        var cell = read.Body.Tables.Single().Rows.Single().Cells.Single();
        Assert.That(cell.Text, Is.EqualTo("outer\ninner\n"));
    }

    [Test]
    public void HyphensAndBreaksAreTheCharactersTheyStandFor()
    {
        using var document = Document.Create();
        var run = document.Body.AddParagraph().AddRun("co");
        run.AppendElement(new W.NoBreakHyphen());
        run.Append("operate");
        run.AppendElement(new W.SoftHyphen());
        run.AppendElement(new W.CarriageReturn());
        run.Append("next");

        using var read = DocumentAssert.Read(document);

        // A non-breaking hyphen and a soft hyphen are stored as elements rather than as text, so a
        // reader that only looks at w:t joins the words either side with nothing between them.
        Assert.That(read.Text, Is.EqualTo("co\u2011operate\u00AD\nnext"));
    }

    [Test]
    public void EverythingHereIsStillAValidDocument()
    {
        using var document = Document.Create();
        document.Body.AddParagraph("text ").AddLink(document, "https://example.org", "link");
        document.Body.AddTable(_ => _.Row("a", "b"));
        DocumentAssert.IsValid(document);
    }
}
