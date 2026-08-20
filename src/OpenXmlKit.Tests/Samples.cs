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
        using var read = DocumentView.Open(bytes);
        Assert.That(read.Body.Tables.Single().Rows.Single().Cells.First().Text, Is.EqualTo("Source"));
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
        using var read = DocumentAssert.Read(document);
        Assert.That(read.Body.Tables.Single().Rows.Single().Cells.Count(), Is.EqualTo(2));
    }

    [Test]
    public void Reading()
    {
        var source = BuildSource();

        #region Reading

        using var document = DocumentView.Open(source);

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
        using var built = Document.Create();
        built.Styles.Add(StyleKind.Table, "Branded", "Branded", style => style.Font.Name = "Georgia");
        built.Styles.Add(StyleKind.Paragraph, "Normal", "Normal", style => style.Font.Name = "Calibri");
        built.Body.Paragraph("cell text", "Normal");

        using var document = DocumentView.Open(built.ToArray());
        var paragraph = document.Body.Paragraphs.Single();
        var run = paragraph.Runs.Single();

        #region ResolvingFormatting

        var font = document.Formatting.FontFor(run, paragraph, tableStyleId: "Branded");

        #endregion

        Assert.That(font.Name, Is.EqualTo("Calibri"));
    }

    [Test]
    public void ReadingContent()
    {
        using var built = Document.Create();
        built.Body.AddParagraph("See ").AddLink(built, "https://example.org/report", "the report");
        built.Body.AddParagraph().AddPageNumber();
        var bullets = built.Numbering.AddBullet();
        var item = built.Body.AddParagraph("a point");
        item.Format.List = bullets.At();
        var source = built.ToArray();

        using var document = DocumentView.Open(source);
        var paragraphs = document.Body.Paragraphs.ToList();

        #region ReadingContent

        // A link, with the relationship resolved back to the address it points at.
        var link = paragraphs[0].Hyperlinks.Single();
        TestContext.Out.WriteLine($"{link.Text} -> {link.Url}");

        // A field, with the instruction and the value Word cached for it.
        var field = paragraphs[1].Fields.Single();
        TestContext.Out.WriteLine($"{field.Code} = {field.Value}");

        // A list paragraph says only which list it is in and how deep; the numbering says what
        // that actually draws.
        var level = document.Numbering.LevelFor(paragraphs[2].List!.Value)!;
        TestContext.Out.WriteLine($"{level.Format} in {level.Font.Name}");

        #endregion

        Assert.That(link.Url, Is.EqualTo("https://example.org/report"));
        Assert.That(field.Code, Is.EqualTo("PAGE"));
        Assert.That(level.Format, Is.EqualTo(NumberFormat.Bullet));
    }

    [Test]
    public void TableStyles()
    {
        using var document = Document.Create();

        #region TableStyles

        document.Styles.Add(
            StyleKind.Table,
            "Branded",
            "Branded",
            style =>
            {
                style.TableFormat.Borders.SetAll(BorderStyle.Single, Length.FromPoints(0.5));

                // What makes a table style a table style: the header row and the banding are
                // conditional blocks over the whole-table formatting above.
                style.Conditional(
                    TableStyleArea.FirstRow,
                    _ =>
                    {
                        _.Font.Bold = true;
                        _.CellFormat.Shading.BackgroundColor = Color.FromRgb(0x223344);
                    });
                style.Conditional(
                    TableStyleArea.Band1Horizontal,
                    _ => _.CellFormat.Shading.BackgroundColor = Color.FromRgb(0xEEEEEE));
            });

        document.Body.AddTable(
            table =>
            {
                table.Style("Branded");
                table.Formatting(format => format.Caption = "Commitment details");
                table.HeaderRow("Source", "Amount");
                table.Row("Grant", "1,200");
            });

        #endregion

        DocumentAssert.IsValid(document);
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
