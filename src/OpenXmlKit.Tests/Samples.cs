/// <summary>
/// The readme's examples, run as tests so they are known to compile and to produce a document Word
/// will open. MarkdownSnippets pulls them into the readme from the region markers.
/// </summary>
[TestFixture]
public class Samples
{
    [Test]
    public void NestedBuilder()
    {
        #region NestedBuilder

        using var document = Document.Create();

        document.Body.AddTable(
            _ => _
                .Style("TableGrid")
                .Width(Width.Percent(100))
                .Row(
                    row => row
                        .Cell(Width.Percent(22), _ => _.AddParagraph(_ => _.Bold("Source")))
                        .Cell(Width.Percent(78), "Budget paper 2")));

        var bytes = document.ToArray();

        #endregion

        DocumentAssert.IsValid(bytes);
        Assert.That(document.Body.Tables.Single().Rows.Single().Cells.First().Text, Is.EqualTo("Source"));
    }

    [Test]
    public void CursorBuilder()
    {
        #region CursorBuilder

        using var document = Document.Create();
        var builder = document.Builder;

        builder.Heading(1, "Delivery update");
        builder.Writeln("The commitment is on schedule.");

        using (builder.PushFormatting())
        {
            builder.Font.Bold = true;
            builder.Writeln("This paragraph is bold.");
        }

        using (builder.Table())
        using (builder.Row())
        {
            builder.InsertCell();
            builder.Write("A");
            builder.InsertCell();
            builder.Write("B");
        }

        #endregion

        DocumentAssert.IsValid(document);
        Assert.That(document.Body.Tables.Single().Rows.Single().Cells.Count(), Is.EqualTo(2));
    }

    [Test]
    public void Reading()
    {
        var source = BuildSource();

        #region Reading

        using var document = Document.Open(new MemoryStream(source));

        foreach (var paragraph in document.Body.Paragraphs)
        {
            TestContext.Out.WriteLine(paragraph.Text);
        }

        #endregion

        Assert.That(document.Body.Paragraphs.First().Text, Is.EqualTo("Delivery update"));
    }

    [Test]
    public void ResolvingFormatting()
    {
        using var document = Document.Create();
        document.Styles.Add(StyleKind.Table, "Branded", "Branded", style => style.Font.Name = "Georgia");
        document.Styles.Add(StyleKind.Paragraph, "Normal", "Normal", style => style.Font.Name = "Calibri");
        var paragraph = document.Body.AddParagraph();
        paragraph.Style("Normal");
        var run = paragraph.AddRun("cell text");

        #region ResolvingFormatting

        var font = document.Formatting.FontFor(run, paragraph, tableStyleId: "Branded");

        #endregion

        Assert.That(font.Name, Is.EqualTo("Calibri"));
    }

    [Test]
    public void EscapeHatch()
    {
        var table = Table.Create().Row("a", "b");

        #region EscapeHatch

        // The SDK's own w:tbl, to hand to code that has not migrated yet.
        var raw = table.ToOpenXml();

        using var document = Document.Create();
        document.Body.AppendElement(raw);

        #endregion

        DocumentAssert.IsValid(document);
    }

    static byte[] BuildSource()
    {
        using var document = Document.Create();
        document.Body.Paragraph("Delivery update", "Heading1");
        document.Body.Paragraph("The commitment is on schedule.");
        return document.ToArray();
    }
}
