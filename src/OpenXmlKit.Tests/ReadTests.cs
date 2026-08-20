[TestFixture]
public class ReadTests
{
    [Test]
    public void RoundTripsContentAndFormatting()
    {
        byte[] bytes;
        using (var written = Document.Create())
        {
            written.Body.Paragraph("Heading", "Heading1");
            written.Body.AddParagraph(
                _ => _
                    .Append("bold ", font => font.Bold = true)
                    .Append("normal"));
            written.Body.AddTable(
                _ => _
                    .Width(Width.Percent(100))
                    .HeaderRow("Name", "Value")
                    .Row("a", "1"));
            bytes = written.ToArray();
        }

        using var read = DocumentView.Open(bytes);

        var paragraphs = read.Body.Paragraphs.ToList();
        Assert.That(paragraphs[0].StyleId, Is.EqualTo("Heading1"));
        Assert.That(paragraphs[1].Text, Is.EqualTo("bold normal"));
        Assert.That(paragraphs[1].Runs.First().Font.Bold.IsOn, Is.True);
        Assert.That(paragraphs[1].Runs.Last().Font.Bold.IsSet, Is.False);

        var table = read.Body.Tables.Single();
        Assert.That(table.Format.Width, Is.EqualTo(Width.Percent(100)));
        Assert.That(table.Rows.First().IsHeader, Is.True);
        Assert.That(table.Rows.Last().Cells.Select(_ => _.Text), Is.EqualTo(new[] { "a", "1" }));
    }

    [Test]
    public void UnitsSurviveTheRoundTrip()
    {
        byte[] bytes;
        using (var written = Document.Create())
        {
            var paragraph = written.Body.AddParagraph("sized");
            paragraph.Format.LeftIndent = Length.FromCentimeters(2);
            paragraph.Format.SpaceAfter = Length.FromPoints(6);
            paragraph.AddRun("text").Font.Size = Length.FromPoints(13.5);
            bytes = written.ToArray();
        }

        using var read = DocumentView.Open(bytes);
        var sized = read.Body.Paragraphs.Single();

        // Twips and half-points are both integers, so a round trip is exact only where the value
        // lands on one. 13.5pt is 27 half-points, and 2cm is 1134 twips.
        Assert.That(sized.Format.SpaceAfter!.Value.TotalPoints, Is.EqualTo(6));
        Assert.That(sized.Format.LeftIndent!.Value.TotalCentimeters, Is.EqualTo(2).Within(0.001));
        Assert.That(sized.Runs.Last().Font.Size!.Value.TotalPoints, Is.EqualTo(13.5));
    }

    [Test]
    public void EffectiveFontWalksTheBasedOnChain()
    {
        using var read = Read(
            document =>
            {
                document.Styles.SetDefaults(font => font.Name = "Calibri");
                document.Styles.Add(
                    StyleKind.Paragraph,
                    "Base",
                    "Base",
                    style =>
                    {
                        style.Font.Size = Length.FromPoints(10);
                        style.Font.Italic = Toggle.On;
                    });
                document.Styles.Add(
                    StyleKind.Paragraph,
                    "Derived",
                    "Derived",
                    style =>
                    {
                        style.BasedOn = "Base";
                        style.Font.Size = Length.FromPoints(14);
                    });
                document.Body.Paragraph("text", "Derived");
            });

        var paragraph = read.Body.Paragraphs.Single();
        var font = read.Formatting.FontFor(paragraph.Runs.Single(), paragraph);

        // Name from the document defaults, italic from the base style, size from the derived one:
        // each level overlays only what it states.
        Assert.That(font.Name, Is.EqualTo("Calibri"));
        Assert.That(font.Italic.IsOn, Is.True);
        Assert.That(font.Size!.Value.TotalPoints, Is.EqualTo(14));
    }

    [Test]
    public void DirectFormattingBeatsEveryStyle()
    {
        using var read = Read(
            document =>
            {
                document.Styles.Add(StyleKind.Paragraph, "Loud", "Loud", style => style.Font.Bold = Toggle.On);
                var paragraph = document.Body.AddParagraph();
                paragraph.Style("Loud");
                paragraph.AddRun("quiet").Font.Bold = Toggle.Off;
            });

        var paragraph = read.Body.Paragraphs.Single();
        var font = read.Formatting.FontFor(paragraph.Runs.Single(), paragraph);

        // The three-state toggle earning its keep: a plain bool could not express this, because
        // its false is indistinguishable from having said nothing.
        Assert.That(font.Bold.IsOff, Is.True);
    }

    [Test]
    public void ParagraphStyleOutranksTableStyle()
    {
        using var read = Read(
            document =>
            {
                document.Styles.Add(StyleKind.Table, "Branded", "Branded", style => style.Font.Name = "Georgia");
                document.Styles.Add(StyleKind.Paragraph, "Normal", "Normal", style => style.Font.Name = "Calibri");
                document.Body.Paragraph("cell text", "Normal");
            });

        var paragraph = read.Body.Paragraphs.Single();
        var font = read.Formatting.FontFor(paragraph.Runs.Single(), paragraph, tableStyleId: "Branded");

        // The precedence rule that catches everybody: a table style sits below the paragraph style,
        // so branding a table through its table style alone loses to whatever Normal says. Naming
        // it here rather than leaving it to be rediscovered is most of why the resolver exists.
        Assert.That(font.Name, Is.EqualTo("Calibri"));
    }

    [Test]
    public void EffectiveFormatCarriesTheStyleThatSuppliedIt()
    {
        using var read = Read(
            document =>
            {
                document.Styles.Add(
                    StyleKind.Paragraph,
                    "Quote",
                    "Quote",
                    style =>
                    {
                        style.ParagraphFormat.LeftIndent = Length.FromInches(0.5);
                        style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
                    });
                var paragraph = document.Body.AddParagraph("quoted");
                paragraph.Style("Quote");
                paragraph.Format.Alignment = ParagraphAlignment.Right;
            });

        var format = read.Formatting.FormatFor(read.Body.Paragraphs.Single());

        Assert.That(format.LeftIndent!.Value.TotalInches, Is.EqualTo(0.5).Within(0.001));
        Assert.That(format.Alignment, Is.EqualTo(ParagraphAlignment.Right));
        Assert.That(format.StyleId, Is.EqualTo("Quote"));
    }

    [Test]
    public void AStyleChainThatLoopsDoesNotHang()
    {
        using var read = Read(
            document =>
            {
                document.Styles.Add(StyleKind.Paragraph, "A", "A", style => style.BasedOn = "B");
                document.Styles.Add(StyleKind.Paragraph, "B", "B", style => style.BasedOn = "A");
                document.Body.Paragraph("text", "A");
            });

        var paragraph = read.Body.Paragraphs.Single();

        // A hand-edited template can carry a cycle, and walking it naively never terminates.
        Assert.That(
            () => read.Formatting.FontFor(paragraph.Runs.Single(), paragraph),
            NUnit.Framework.Throws.Nothing);
    }

    [Test]
    public void SectionsAndPageSetupAreReadBack()
    {
        using var read = Read(
            document =>
            {
                document.Body.Section.PageSetup.SetA4();
                document.Body.Paragraph("Portrait.");
                document.Body.AddSection().PageSetup.SetA4(PageOrientation.Landscape);
                document.Body.Paragraph("Landscape.");
            });

        var sections = read.Sections.ToList();

        Assert.That(sections, Has.Count.EqualTo(2));
        Assert.That(sections[0].PageSetup.Orientation, Is.EqualTo(PageOrientation.Portrait));
        Assert.That(sections[1].PageSetup.Orientation, Is.EqualTo(PageOrientation.Landscape));
    }

    [Test]
    public void HeadersAreReachableFromTheirSection()
    {
        using var read = Read(
            document =>
            {
                document.Body.Section.AddHeader().Paragraph("OFFICIAL");
                document.Body.Paragraph("Body.");
            });

        var header = read.Sections.Last().Headers.Single();

        Assert.That(header.Key, Is.EqualTo(HeaderFooterKind.Default));
        Assert.That(header.Value.Text, Is.EqualTo("OFFICIAL"));
    }

    [Test]
    public void StylesAreReadBackWithTheirInheritance()
    {
        using var read = Read(document => document.Styles.EnsureBuiltIn(BuiltInStyle.TableGrid));

        var grid = read.Styles[BuiltInStyle.TableGrid];

        Assert.That(grid, Is.Not.Null);
        Assert.That(grid!.Value.BasedOn, Is.EqualTo("TableNormal"));
        Assert.That(grid.Value.Kind, Is.EqualTo(StyleKind.Table));
    }

    [Test]
    public void BookmarksAreReadBack()
    {
        using var read = Read(
            document => document.Body.AddParagraph()
                .AddBookmark(document, "du3", _ => _.Append("The third update")));

        Assert.That(read.Body.Paragraphs.Single().BookmarkNames, Is.EqualTo(new[] { "du3" }));
    }

    [Test]
    public void OpeningADocumentContinuesItsNumberingIdsRatherThanColliding()
    {
        byte[] bytes;
        int firstId;
        using (var written = Document.Create())
        {
            var list = written.Numbering.AddNumbered();
            firstId = list.NumberingId;
            written.Builder.ListItem(list).Writeln("one");
            bytes = written.ToArray();
        }

        // An expandable stream, because appending means writing back: a MemoryStream built over a
        // byte array is fixed-size and fails on the first write with an obscure message.
        var editable = new MemoryStream();
        editable.Write(bytes, 0, bytes.Length);
        editable.Position = 0;

        using var reopened = Document.OpenForAppend(editable);
        var second = reopened.Numbering.AddBullet();

        // Continuing rather than restarting is what stops a second call quietly taking over the
        // numbering of the first.
        Assert.That(second.NumberingId, Is.GreaterThan(firstId));
    }

    static DocumentView Read(Action<Document> build)
    {
        using var document = Document.Create();
        build(document);
        return DocumentView.Open(document.ToArray());
    }
}
