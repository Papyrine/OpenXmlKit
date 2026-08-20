[TestFixture]
public class DeterminismTests
{
    [Test]
    public void TheSameContentProducesTheSamePartXml()
    {
        var first = BuildParts();
        var second = BuildParts();

        // What this library controls, and all of it: the same calls produce byte-identical part
        // XML. Relationship ids are pinned rather than generated, and nothing in the emitters
        // depends on iteration order or on the clock.
        Assert.That(second, Is.EqualTo(first));
    }

    [Test]
    public void ThePackageItselfIsNotByteIdenticalWithoutDeterministicPackaging()
    {
        var first = Build();
        var second = Build();

        // Worth pinning as a known limit rather than leaving to be discovered. The zip entries
        // carry their own modification timestamps, written by the packaging layer beneath the SDK,
        // so two runs a second apart differ in bytes while being identical in content.
        //
        // A consumer that needs byte equality — the estate paginates documents itself and writes
        // page numbers in, so it does — adds DeterministicIoPackaging, which replaces that layer.
        // That is deliberately the consumer's dependency rather than this package's: it changes how
        // every document in the process is written, which is not a decision a wrapper should make
        // on their behalf.
        Assert.That(second, Is.Not.EqualTo(first));
    }

    [Test]
    public void RelationshipIdsArePinnedRatherThanGenerated()
    {
        using var document = Document.Create();
        document.Styles.EnsureBuiltIn(BuiltInStyle.Normal);
        document.Numbering.AddBullet();
        document.Body.AddParagraph().AddImage(document, SamplePng.Create(8, 8));

        using var stream = new MemoryStream(document.ToArray());
        using var package = WordprocessingDocument.Open(stream, false);
        var main = package.MainDocumentPart!;

        Assert.That(main.GetIdOfPart(main.StyleDefinitionsPart!), Is.EqualTo("rStyles"));
        Assert.That(main.GetIdOfPart(main.NumberingDefinitionsPart!), Is.EqualTo("rNumbering"));
        Assert.That(main.GetIdOfPart(main.ImageParts.Single()), Is.EqualTo("rImage1"));
    }

    static string BuildParts()
    {
        using var stream = new MemoryStream(Build());
        using var package = WordprocessingDocument.Open(stream, false);
        var main = package.MainDocumentPart!;
        return string.Join(
            "\n",
            main.Document!.OuterXml,
            main.StyleDefinitionsPart?.Styles?.OuterXml ?? "",
            main.NumberingDefinitionsPart?.Numbering?.OuterXml ?? "");
    }

    static byte[] Build()
    {
        using var document = Document.Create();
        // Fixed rather than "now", because the created and modified timestamps are otherwise the
        // moment of generation and no two runs can agree.
        document.Properties.Created = new(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        document.Properties.Modified = document.Properties.Created;
        document.Properties.Title = "Determinism";

        document.Styles.EnsureBuiltIn(BuiltInStyle.TableGrid);
        document.Body.Paragraph("Heading", "Heading1");
        document.Body.AddTable(
            _ => _
                .Style("TableGrid")
                .Width(Width.Percent(100))
                .HeaderRow("Name", "Value")
                .Row("a", "1"));
        return document.ToArray();
    }
}
