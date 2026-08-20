/// <summary>
/// Everything the build API can write, read back.
/// </summary>
/// <remarks>
/// The read half is a separate type hierarchy from the build half, which makes it possible for the
/// two to drift: a document could carry a link, a picture or a footnote that nothing on the read
/// side could reach. Each test here writes with one API and asserts with the other, so a gap
/// between them fails rather than going unnoticed.
/// </remarks>
[TestFixture]
public class ReadContentTests
{
    [Test]
    public void LinksAreReadable()
    {
        using var document = Document.Create();
        document.Body.AddParagraph().AddLink(document, "https://example.org/report", "the report");
        document.Body.AddParagraph().AddAnchorLink(document, "Summary", "see the summary");

        using var read = DocumentAssert.Read(document);
        var paragraphs = read.Body.Paragraphs.ToList();

        var external = paragraphs[0].Hyperlinks.Single();
        Assert.That(external.Url, Is.EqualTo("https://example.org/report"));
        Assert.That(external.Text, Is.EqualTo("the report"));
        Assert.That(external.Anchor, Is.Null);

        var internalLink = paragraphs[1].Hyperlinks.Single();
        Assert.That(internalLink.Anchor, Is.EqualTo("Summary"));
        Assert.That(internalLink.Url, Is.Null);
        Assert.That(internalLink.Text, Is.EqualTo("see the summary"));
    }

    [Test]
    public void PicturesAreReadable()
    {
        var png = SamplePng.Create(40, 20);

        using var document = Document.Create();
        document.Body.AddParagraph()
            .AddImage(document, png, width: Length.FromInches(2), description: "A chart");

        using var read = DocumentAssert.Read(document);
        var image = read.Body.Paragraphs.Single().Images.Single();

        Assert.That(image.Description, Is.EqualTo("A chart"));
        Assert.That(image.IsFloating, Is.False);
        Assert.That(image.Width.TotalInches, Is.EqualTo(2).Within(0.001));
        // Height was not given, so it followed the image's own 2:1 aspect ratio.
        Assert.That(image.Height.TotalInches, Is.EqualTo(1).Within(0.001));
        Assert.That(image.ContentType, Does.Contain("png"));
        Assert.That(image.GetBytes(), Is.EqualTo(png));
    }

    [Test]
    public void AFloatedPictureSaysSo()
    {
        using var document = Document.Create();
        document.Body.AddParagraph()
            .AddImage(document, SamplePng.Create(10, 10), wrap: ImageWrap.Right);

        using var read = DocumentAssert.Read(document);
        Assert.That(read.Body.Paragraphs.Single().Images.Single().IsFloating, Is.True);
    }

    [Test]
    public void FieldsAreReadableWithTheirCachedValue()
    {
        using var document = Document.Create();
        document.Body.AddParagraph().AddPageNumber();
        document.Body.AddParagraph().AddPageReference("Summary", "12");

        using var read = DocumentAssert.Read(document);
        var paragraphs = read.Body.Paragraphs.ToList();

        Assert.That(paragraphs[0].Fields.Single().Code, Is.EqualTo("PAGE"));

        var reference = paragraphs[1].Fields.Single();
        Assert.That(reference.Code, Is.EqualTo(@"PAGEREF Summary \h"));
        Assert.That(reference.Value, Is.EqualTo("12"));
    }

    [Test]
    public void FootnotesAreReadable()
    {
        using var document = Document.Create();
        document.Body.AddParagraph("Revenue rose.")
            .AddFootnote(document, _ => _.AddParagraph("Against the revised schedule."));

        using var read = DocumentAssert.Read(document);
        var footnote = read.Footnotes.Single();

        Assert.That(footnote.Text, Does.Contain("Against the revised schedule."));

        // The separator and continuation-separator notes Word requires are machinery, not content.
        Assert.That(read.Footnotes.Count(), Is.EqualTo(1));

        var reference = read.Body.Paragraphs.Single().FootnoteReferences.Single();
        Assert.That(reference, Is.EqualTo(footnote.Id));
    }

    [Test]
    public void AListResolvesToTheMarkerItDraws()
    {
        using var document = Document.Create();
        var numbered = document.Numbering.AddNumbered();
        var paragraph = document.Body.AddParagraph("first");
        paragraph.Format.List = numbered.At();

        using var read = DocumentAssert.Read(document);
        var membership = read.Body.Paragraphs.Single().List!.Value;

        // Reading the membership alone says only "list n, level 0" — which is not enough to know
        // whether that renders as "1." or as a bullet.
        var level = read.Numbering.LevelFor(membership)!;
        Assert.That(level.Format, Is.EqualTo(NumberFormat.Decimal));
        Assert.That(level.Text, Is.EqualTo("%1."));
        Assert.That(level.StartAt, Is.EqualTo(1));
    }

    [Test]
    public void ABulletResolvesToItsGlyphAndFont()
    {
        using var document = Document.Create();
        var bullets = document.Numbering.AddBullet();
        var paragraph = document.Body.AddParagraph("point");
        paragraph.Format.List = bullets.At();

        using var read = DocumentAssert.Read(document);
        var level = read.Numbering.LevelFor(read.Body.Paragraphs.Single().List!.Value)!;

        Assert.That(level.Format, Is.EqualTo(NumberFormat.Bullet));
        // The glyph comes from the font named on the level rather than from the paragraph's own.
        Assert.That(level.Font.Name, Is.EqualTo("Symbol"));
        Assert.That(level.Hanging.TotalInches, Is.EqualTo(0.25).Within(0.001));
    }

    [Test]
    public void EveryBlockContainerIsReachable()
    {
        using var document = Document.Create();
        document.Body.AddParagraph("body text")
            .AddFootnote(document, _ => _.AddParagraph("note text"));
        var section = document.Body.Section;
        section.AddHeader(HeaderFooterKind.Default).AddParagraph("header text");
        section.AddFooter(HeaderFooterKind.Default).AddParagraph("footer text");

        using var read = DocumentAssert.Read(document);
        var text = read.Containers.Select(_ => _.Value.Text).ToList();

        // Body alone reports the letterhead and every footnote as absent.
        Assert.That(text.Any(_ => _.Contains("body text")), Is.True);
        Assert.That(text.Any(_ => _.Contains("header text")), Is.True);
        Assert.That(text.Any(_ => _.Contains("footer text")), Is.True);
        Assert.That(text.Any(_ => _.Contains("note text")), Is.True);
    }

    [Test]
    public void AStyleTableFormatIsReadable()
    {
        using var document = Document.Create();
        document.Styles.Add(
            StyleKind.Table,
            "Branded",
            "Branded",
            style =>
            {
                style.IsQuickStyle = true;
                style.TableFormat.Borders.SetAll(BorderStyle.Single, Length.FromPoints(0.5));
                style.TableFormat.SetDefaultMargins(Length.FromPoints(4), Length.FromPoints(2));
            });

        using var read = DocumentAssert.Read(document);
        var style = read.Styles["Branded"]!.Value;

        Assert.That(style.IsQuickStyle, Is.True);
        Assert.That(style.TableFormat.Borders.Top.Style, Is.EqualTo(BorderStyle.Single));
        Assert.That(style.TableFormat.DefaultLeftMargin!.Value.TotalPoints, Is.EqualTo(4));
    }
}
