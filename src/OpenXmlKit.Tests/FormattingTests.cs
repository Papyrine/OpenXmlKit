/// <summary>
/// Formatting the estate needs that the first cut of the format objects did not carry.
/// </summary>
[TestFixture]
public class FormattingTests
{
    [Test]
    public void ProofingLanguageRoundTrips()
    {
        using var document = Document.Create();
        var run = document.Body.AddParagraph().AddRun("colour");
        run.Font.Language = "en-AU";

        var bytes = DocumentAssert.IsValid(document);

        using var read = DocumentView.Open(bytes);
        Assert.That(read.Body.Paragraphs.Single().Runs.Single().Font.Language, Is.EqualTo("en-AU"));
    }

    [Test]
    public void ProofingLanguageCarriesTheScriptVariantsSeparately()
    {
        using var document = Document.Create();
        var run = document.Body.AddParagraph().AddRun("text");
        run.Font.Language = "en-AU";
        run.Font.LanguageEastAsia = "ja-JP";
        run.Font.LanguageComplexScript = "ar-SA";

        var bytes = DocumentAssert.IsValid(document);

        using var read = DocumentView.Open(bytes);
        var font = read.Body.Paragraphs.Single().Runs.Single().Font;
        Assert.That(font.Language, Is.EqualTo("en-AU"));
        Assert.That(font.LanguageEastAsia, Is.EqualTo("ja-JP"));
        Assert.That(font.LanguageComplexScript, Is.EqualTo("ar-SA"));
    }

    [Test]
    public void LanguageOnAStyleResolvesOntoARun()
    {
        using var document = Document.Create();
        document.Styles.SetDefaults(font => font.Language = "en-AU");
        document.Body.AddParagraph("text");

        var bytes = DocumentAssert.IsValid(document);

        using var read = DocumentView.Open(bytes);
        var paragraph = read.Body.Paragraphs.Single();
        var resolved = read.Formatting.FontFor(paragraph.Runs.Single(), paragraph, null);
        Assert.That(resolved.Language, Is.EqualTo("en-AU"));
    }

    [Test]
    public void ARaggedRowStatesTheGridItSkips()
    {
        using var document = Document.Create();
        document.Body.AddTable(
            table =>
            {
                table.Columns(Length.FromInches(1), Length.FromInches(1), Length.FromInches(1));
                table.Row("a", "b", "c");
                table.AddRow(
                    row =>
                    {
                        // Every row shares the table's one grid, so a row that starts part-way
                        // across states the gap rather than carrying placeholder cells.
                        row.Format.GridBefore = 2;
                        row.Format.WidthBefore = Width.FromPoints(144);
                        row.Cell("only");
                    });
            });

        var bytes = DocumentAssert.IsValid(document);

        using var read = DocumentView.Open(bytes);
        var ragged = read.Body.Tables.Single().Rows.Last();

        Assert.That(ragged.Cells.Count(), Is.EqualTo(1));
        Assert.That(ragged.Format.GridBefore, Is.EqualTo(2));
        Assert.That(ragged.Format.WidthBefore, Is.EqualTo(Width.FromPoints(144)));
    }

    [Test]
    public void GridAfterRoundTrips()
    {
        using var document = Document.Create();
        document.Body.AddTable(
            table =>
            {
                table.Row("a", "b");
                table.AddRow(
                    row =>
                    {
                        row.Format.GridAfter = 1;
                        row.Format.WidthAfter = Width.FromPoints(72);
                        row.Cell("only");
                    });
            });

        var bytes = DocumentAssert.IsValid(document);

        using var read = DocumentView.Open(bytes);
        var row = read.Body.Tables.Single().Rows.Last();
        Assert.That(row.Format.GridAfter, Is.EqualTo(1));
        Assert.That(row.Format.WidthAfter, Is.EqualTo(Width.FromPoints(72)));
    }

    [Test]
    public void ATableCarriesItsAlternativeText()
    {
        using var document = Document.Create();
        document.Body.AddTable(
            table =>
            {
                table.Formatting(
                    format =>
                    {
                        format.Caption = "Commitment details";
                        format.Description = "Source, amount and date for each commitment.";
                    });
                table.Row("a", "b");
            });

        var bytes = DocumentAssert.IsValid(document);

        using var read = DocumentView.Open(bytes);
        var format = read.Body.Tables.Single().Format;
        Assert.That(format.Caption, Is.EqualTo("Commitment details"));
        Assert.That(format.Description, Is.EqualTo("Source, amount and date for each commitment."));
    }
}
