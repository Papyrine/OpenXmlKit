/// <summary>
/// Conditional table formatting — what makes a table style more than a set of borders.
/// </summary>
[TestFixture]
public class TableStyleTests
{
    [Test]
    public void ConditionalBlocksAreWrittenAndReadBack()
    {
        using var document = Document.Create();
        document.Styles.Add(
            StyleKind.Table,
            "Branded",
            "Branded",
            style =>
            {
                style.TableFormat.Borders.SetAll(BorderStyle.Single, Length.FromPoints(0.5));
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

        var bytes = DocumentAssert.IsValid(document);

        using var read = DocumentView.Open(bytes);
        var conditionals = read.Styles["Branded"]!.Value.ConditionalFormats.ToList();

        Assert.That(conditionals.Select(_ => _.Area), Is.EqualTo([TableStyleArea.FirstRow, TableStyleArea.Band1Horizontal]));
        Assert.That(conditionals[0].Font.Bold.IsOn, Is.True);
        Assert.That(conditionals[0].CellFormat.Shading.BackgroundColor, Is.EqualTo(Color.FromRgb(0x223344)));
        Assert.That(conditionals[1].CellFormat.Shading.BackgroundColor, Is.EqualTo(Color.FromRgb(0xEEEEEE)));
    }

    [Test]
    public void AskingForTheSameAreaTwiceConfiguresTheSameBlock()
    {
        using var document = Document.Create();
        var style = document.Styles.Add(StyleKind.Table, "Branded");
        style.Conditional(TableStyleArea.FirstRow).Font.Bold = true;
        style.Conditional(TableStyleArea.FirstRow).Font.AllCaps = true;

        var bytes = DocumentAssert.IsValid(document);
        using var read = DocumentView.Open(bytes);

        var conditional = read.Styles["Branded"]!.Value.ConditionalFormats.Single();
        Assert.That(conditional.Font.Bold.IsOn, Is.True);
        Assert.That(conditional.Font.AllCaps.IsOn, Is.True);
    }

    [Test]
    public void TheCornerCellsMapToWordsCompassNames()
    {
        using var document = Document.Create();
        var style = document.Styles.Add(StyleKind.Table, "Corners");
        style.Conditional(TableStyleArea.TopLeftCell).Font.Bold = true;
        style.Conditional(TableStyleArea.BottomRightCell).Font.Italic = true;

        var xml = DocumentAssert.PartXml(document, "styles")!;
        Assert.That(xml, Does.Contain("w:type=\"nwCell\""));
        Assert.That(xml, Does.Contain("w:type=\"seCell\""));

        using var read = DocumentView.Open(document.ToArray());
        Assert.That(
            read.Styles["Corners"]!.Value.ConditionalFormats.Select(_ => _.Area),
            Is.EqualTo([TableStyleArea.TopLeftCell, TableStyleArea.BottomRightCell]));
    }

    [Test]
    public void FormattingTheSchemaForbidsInAnOverrideThrows()
    {
        // Deliberately not disposed: writing the styles part is what detects this, and disposing
        // saves, so a using block would raise the same error a second time out of the teardown.
        var document = Document.Create();
        var style = document.Styles.Add(StyleKind.Table, "Branded");

        // A conditional block is CT_TcPrStyleOverride, which has no tcW: a width belongs to the
        // cell using the style, not to the style. Dropping it silently would leave a style that
        // does less than it says, so it is rejected instead.
        style.Conditional(TableStyleArea.FirstRow).CellFormat.Width = Width.Percent(50);

        var exception = Assert.Throws<InvalidOperationException>(() => document.ToArray());
        Assert.That(exception!.Message, Does.Contain("tcW"));
    }

    [Test]
    public void ARunOverrideCannotNameAnotherStyle()
    {
        var document = Document.Create();
        var style = document.Styles.Add(StyleKind.Table, "Branded");
        style.Conditional(TableStyleArea.FirstRow).Font.StyleId = "Emphasis";

        var exception = Assert.Throws<InvalidOperationException>(() => document.ToArray());
        Assert.That(exception!.Message, Does.Contain("rStyle"));
    }

    [Test]
    public void AStyleWithConditionalsKeepsItsTablePropertiesFirst()
    {
        using var document = Document.Create();
        var style = document.Styles.Add(StyleKind.Table, "Branded");
        style.Conditional(TableStyleArea.FirstRow).Font.Bold = true;
        style.TableFormat.Borders.SetAll(BorderStyle.Single);

        // tblPr precedes tblStylePr in CT_Style's sequence, and the two are written by separate
        // passes, so the order between them is the thing worth pinning.
        var xml = DocumentAssert.PartXml(document, "styles")!;
        Assert.That(xml.IndexOf("<w:tblPr>", StringComparison.Ordinal), Is.LessThan(xml.IndexOf("<w:tblStylePr", StringComparison.Ordinal)));
        DocumentAssert.IsValid(document);
    }
}
