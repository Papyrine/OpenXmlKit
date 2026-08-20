[TestFixture]
public class BuildTests
{
    [Test]
    public void EmptyDocumentIsValid()
    {
        using var document = Document.Create();
        DocumentAssert.IsValid(document);
    }

    [Test]
    public void ParagraphsAndRuns()
    {
        using var document = Document.Create();
        document.Body
            .Paragraph("Plain text.")
            .AddParagraph(
                _ => _
                    .Append("Bold ", font => font.Bold = true)
                    .Append("and italic ", font => font.Italic = true)
                    .Append("and coloured.", font => font.Color = Color.Parse("#C00000")));

        DocumentAssert.IsValid(document);
        using var read = DocumentAssert.Read(document);
        Assert.That(read.Body.Text, Is.EqualTo("Plain text.\nBold and italic and coloured."));
    }

    [Test]
    public void NewlinesBecomeBreaksAndTabsBecomeTabs()
    {
        using var document = Document.Create();
        document.Body.Paragraph("first\nsecond\tindented");

        var xml = DocumentAssert.MainPartXml(document);
        DocumentAssert.IsValid(document);

        // Word ignores both characters inside a text element, so a value carrying its own line
        // breaks would otherwise run together on one line.
        Assert.That(xml, Does.Contain("<w:br />"));
        Assert.That(xml, Does.Contain("<w:tab />"));
        using var read = DocumentAssert.Read(document);
        Assert.That(read.Body.Text, Is.EqualTo("first\nsecond\tindented"));
    }

    [Test]
    public void WhitespaceIsPreserved()
    {
        using var document = Document.Create();
        document.Body.Paragraph("  leading and trailing  ");

        DocumentAssert.IsValid(document);
        Assert.That(DocumentAssert.MainPartXml(document), Does.Contain("xml:space=\"preserve\""));
    }

    [Test]
    public void ToggleOffCancelsAnInheritedStyle()
    {
        using var document = Document.Create();
        var paragraph = document.Body.AddParagraph();
        paragraph.AddRun("not bold").Font.Bold = Toggle.Off;

        var xml = DocumentAssert.MainPartXml(document);
        DocumentAssert.IsValid(document);

        // The point of the three-state toggle: an explicit off has to reach the file, or a run in a
        // bold style can never be un-bolded.
        Assert.That(xml, Does.Contain("<w:b w:val=\"false\" />"));
    }

    [Test]
    public void ToggleInheritWritesNothing()
    {
        using var document = Document.Create();
        document.Body.AddParagraph().AddRun("plain");

        var xml = DocumentAssert.MainPartXml(document);
        DocumentAssert.IsValid(document);

        Assert.That(xml, Does.Not.Contain("<w:b "));
        Assert.That(xml, Does.Not.Contain("<w:b/"));
        Assert.That(xml, Does.Not.Contain("<w:rPr"));
    }

    [Test]
    public void FontSizeIsWrittenInHalfPoints()
    {
        using var document = Document.Create();
        document.Body.AddParagraph().AddRun("twelve point").Font.Size = 12;

        DocumentAssert.IsValid(document);
        Assert.That(DocumentAssert.MainPartXml(document), Does.Contain("<w:sz w:val=\"24\" />"));
    }

    [Test]
    public void BorderWidthIsWrittenInEighthsOfAPoint()
    {
        using var document = Document.Create();
        document.Body.AddTable(
            _ => _
                .Borders(BorderStyle.Single, Length.FromPoints(0.5))
                .Row("one"));

        DocumentAssert.IsValid(document);
        // Half a point is four eighths — the same attribute name as font size, a different unit.
        Assert.That(DocumentAssert.MainPartXml(document), Does.Contain("w:sz=\"4\""));
    }

    [Test]
    public void TableWithPercentageWidths()
    {
        using var document = Document.Create();
        document.Body.AddTable(
            _ => _
                .Width(Width.Percent(100))
                .Row(
                    row => row
                        .Cell(Width.Percent(22), "Source")
                        .Cell(Width.Percent(78), "Budget paper 2")));

        var xml = DocumentAssert.MainPartXml(document);
        DocumentAssert.IsValid(document);

        // Word encodes a percentage as fiftieths of a percent.
        Assert.That(xml, Does.Contain("w:w=\"5000\" w:type=\"pct\""));
        Assert.That(xml, Does.Contain("w:w=\"1100\" w:type=\"pct\""));
    }

    [Test]
    public void TableGridIsWrittenEvenWithoutDeclaredWidths()
    {
        using var document = Document.Create();
        document.Body.AddTable(_ => _.Row("a", "b", "c"));

        var xml = DocumentAssert.MainPartXml(document);
        DocumentAssert.IsValid(document);

        // A grid narrower than a row makes Word drop the surplus cells, so one column per cell is
        // written whether or not any widths were declared.
        Assert.That(Regex.Matches(xml, "<w:gridCol").Count, Is.EqualTo(3));
    }

    [Test]
    public void TableCellShadingAndSpans()
    {
        using var document = Document.Create();
        document.Body.AddTable(
            _ => _
                .Row(
                    row => row
                        .AddCell(_ => _.Background(Color.Parse("#E0E8F2")).Paragraph("shaded"))
                        .AddCell(_ => _.ColumnSpan(2).Paragraph("wide"))));

        var xml = DocumentAssert.MainPartXml(document);
        DocumentAssert.IsValid(document);

        Assert.That(xml, Does.Contain("w:fill=\"E0E8F2\""));
        Assert.That(xml, Does.Contain("<w:gridSpan w:val=\"2\" />"));
    }

    [Test]
    public void HeaderRowRepeatsAcrossPages()
    {
        using var document = Document.Create();
        document.Body.AddTable(
            _ => _
                .HeaderRow("Name", "Value")
                .Row("a", "1"));

        var xml = DocumentAssert.MainPartXml(document);
        DocumentAssert.IsValid(document);

        Assert.That(xml, Does.Contain("<w:tblHeader"));
    }

    [Test]
    public void BuiltInStylesAreSeededWithTheirDependencies()
    {
        using var document = Document.Create();
        document.Styles.EnsureBuiltIn(BuiltInStyle.TableGrid);
        document.Body.AddTable(_ => _.Style("TableGrid").Row("a"));

        var styles = DocumentAssert.PartXml(document, "styles");
        DocumentAssert.IsValid(document);

        Assert.That(styles, Is.Not.Null);
        // TableGrid draws borders and nothing else; the cell padding comes from TableNormal, so
        // asking for one has to bring the other.
        Assert.That(styles, Does.Contain("w:styleId=\"TableGrid\""));
        Assert.That(styles, Does.Contain("w:styleId=\"TableNormal\""));
    }

    [Test]
    public void EnsureBuiltInLeavesAnExistingDefinitionAlone()
    {
        using var document = Document.Create();
        var mine = document.Styles.Add(StyleKind.Table, "TableGrid", "My Table Grid");
        document.Styles.EnsureBuiltIn(BuiltInStyle.TableGrid);

        Assert.That(document.Styles["TableGrid"]!.Name, Is.EqualTo("My Table Grid"));
        Assert.That(document.Styles["TableGrid"], Is.SameAs(mine));
        DocumentAssert.IsValid(document);
    }

    [Test]
    public void HeadingsCarryOutlineLevels()
    {
        using var document = Document.Create();
        document.Builder
            .Heading(1, "Delivery update")
            .Writeln("Body text.");

        var styles = DocumentAssert.PartXml(document, "styles");
        DocumentAssert.IsValid(document);

        Assert.That(DocumentAssert.MainPartXml(document), Does.Contain("<w:pStyle w:val=\"Heading1\" />"));
        // Level 0, because Word counts outline levels from zero. This is what puts a heading in the
        // navigation pane and in a generated table of contents.
        Assert.That(styles, Does.Contain("<w:outlineLvl w:val=\"0\" />"));
    }

    [Test]
    public void NumberedAndBulletedLists()
    {
        using var document = Document.Create();
        var numbered = document.Numbering.AddNumbered();
        var bullets = document.Numbering.AddBullet();

        var builder = document.Builder;
        builder.ListItem(numbered);
        builder.Writeln("first");
        builder.Writeln("second");
        builder.ListItem(bullets);
        builder.Writeln("a bullet");
        builder.EndList();
        builder.Writeln("back to prose");

        var numbering = DocumentAssert.PartXml(document, "numbering");
        DocumentAssert.IsValid(document);

        Assert.That(numbering, Is.Not.Null);
        Assert.That(numbered.NumberingId, Is.Not.EqualTo(bullets.NumberingId));
        Assert.That(DocumentAssert.MainPartXml(document), Does.Contain("<w:numPr>"));
    }

    [Test]
    public void RestartedListSharesAppearanceButNotCounting()
    {
        using var document = Document.Create();
        var first = document.Numbering.AddNumbered();
        var second = document.Numbering.Restart(first);

        document.Builder.ListItem(first).Writeln("one");
        document.Builder.ListItem(second).Writeln("one again");

        DocumentAssert.IsValid(document);
        var numbering = DocumentAssert.PartXml(document, "numbering")!;

        Assert.That(second.NumberingId, Is.Not.EqualTo(first.NumberingId));
        // One appearance, two instances: the numbering belongs to the instance, so the second list
        // counts from the start while looking identical.
        Assert.That(Regex.Matches(numbering, "<w:abstractNum ").Count, Is.EqualTo(1));
        Assert.That(Regex.Matches(numbering, "<w:num ").Count, Is.EqualTo(2));
    }

    [Test]
    public void PageSetupAndLandscapeSection()
    {
        using var document = Document.Create();
        document.Body.Section.PageSetup.SetA4();
        document.Body.Section.PageSetup.SetMargins(Length.FromCentimeters(2));
        document.Body.Paragraph("Portrait.");

        var landscape = document.Body.AddSection();
        landscape.PageSetup.SetA4(PageOrientation.Landscape);
        document.Body.Paragraph("Landscape.");

        var xml = DocumentAssert.MainPartXml(document);
        DocumentAssert.IsValid(document);

        Assert.That(xml, Does.Contain("w:orient=\"landscape\""));
        // A4 portrait is 11906 x 16838 twips; landscape is the same numbers the other way round,
        // because Word takes the measurements literally rather than reading the flag.
        Assert.That(xml, Does.Contain("w:w=\"11906\" w:h=\"16838\""));
        Assert.That(xml, Does.Contain("w:w=\"16838\" w:h=\"11906\""));
    }

    [Test]
    public void HeaderReferencesLeadSectionProperties()
    {
        using var document = Document.Create();
        document.Body.Section.PageSetup.SetA4();
        document.Body.Section.AddHeader().Paragraph("OFFICIAL");
        document.Body.Section.AddFooter().Paragraph("Page");
        document.Body.Paragraph("Body.");

        var xml = DocumentAssert.MainPartXml(document);
        DocumentAssert.IsValid(document);

        // Word repairs a document whose sectPr states page size before its header and footer
        // references, and repairing drops the headers.
        var sectionStart = xml.IndexOf("<w:sectPr", StringComparison.Ordinal);
        var headerAt = xml.IndexOf("<w:headerReference", sectionStart, StringComparison.Ordinal);
        var pageSizeAt = xml.IndexOf("<w:pgSz", sectionStart, StringComparison.Ordinal);
        Assert.That(headerAt, Is.GreaterThan(-1));
        Assert.That(headerAt, Is.LessThan(pageSizeAt));
    }

    [Test]
    public void FirstPageHeaderSetsTheFlagThatMakesItShow()
    {
        using var document = Document.Create();
        document.Body.Section.AddHeader(HeaderFooterKind.First).Paragraph("Cover");
        document.Body.Paragraph("Body.");

        DocumentAssert.IsValid(document);
        // Without titlePg the header is stored and never rendered, which is a silent failure worth
        // not leaving to the caller.
        Assert.That(DocumentAssert.MainPartXml(document), Does.Contain("<w:titlePg />"));
    }

    [Test]
    public void CursorBuilderAndNestedBuilderProduceTheSameTable()
    {
        using var nested = Document.Create();
        nested.Body.Append(
            Table.Create()
                .Width(Width.Percent(100))
                .Row(_ => _.Cell(Width.Percent(50), "a").Cell(Width.Percent(50), "b")));

        using var cursor = Document.Create();
        var builder = cursor.Builder;
        builder.CellFormat.Width = Width.Percent(50);
        using (builder.Table())
        {
            builder.TableFormat.Width = Width.Percent(100);
            using (builder.Row())
            {
                builder.InsertCell();
                builder.Write("a");
                builder.InsertCell();
                builder.Write("b");
            }
        }

        DocumentAssert.IsValid(nested);
        DocumentAssert.IsValid(cursor);

        using var readNested = DocumentAssert.Read(nested);
        using var readCursor = DocumentAssert.Read(cursor);
        Assert.That(
            readCursor.Body.Tables.Single().Rows.Single().Cells.Select(_ => _.Text),
            Is.EqualTo(readNested.Body.Tables.Single().Rows.Single().Cells.Select(_ => _.Text)));
    }

    [Test]
    public void PushFormattingRestoresEverythingItSaved()
    {
        using var document = Document.Create();
        var builder = document.Builder;
        builder.Font.Bold = true;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        using (builder.PushFormatting())
        {
            builder.Font.Bold = false;
            builder.Font.Italic = true;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Writeln("inner");
        }

        builder.Writeln("outer");
        DocumentAssert.IsValid(document);

        Assert.That(builder.Font.Bold.IsOn, Is.True);
        Assert.That(builder.Font.Italic.IsSet, Is.False);
        Assert.That(builder.ParagraphFormat.Alignment, Is.EqualTo(ParagraphAlignment.Center));
    }

    [Test]
    public void SymbolCharacterDrawsACheckbox()
    {
        using var document = Document.Create();
        document.Body.AddParagraph().AddRun().AppendSymbol("Wingdings", '');

        DocumentAssert.IsValid(document);
        Assert.That(DocumentAssert.MainPartXml(document), Does.Contain("w:font=\"Wingdings\""));
    }

    [Test]
    public void RawElementsCanBeAppended()
    {
        using var document = Document.Create();
        // The escape hatch: anything the library does not model is still reachable.
        document.Body.AddParagraph()
            .AppendElement(new W.Run(new W.Text("raw")));

        DocumentAssert.IsValid(document);
        using var read = DocumentAssert.Read(document);
        Assert.That(read.Body.Text, Is.EqualTo("raw"));
    }

    [Test]
    public void BuiltTableCanBeHandedOverAsARawElement()
    {
        var table = Table.Create()
            .Width(Width.Percent(100))
            .Row("a", "b")
            .ToOpenXml();

        using var document = Document.Create();
        document.Body.AppendElement(table);

        DocumentAssert.IsValid(document);
        using var read = DocumentAssert.Read(document);
        Assert.That(read.Body.Tables.Single().Rows.Single().Cells.Count(), Is.EqualTo(2));
    }
}
