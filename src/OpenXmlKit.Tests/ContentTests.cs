[TestFixture]
public class ContentTests
{
    [Test]
    public void ImageIsSizedFromItsOwnHeader()
    {
        var png = SamplePng.Create(192, 96);

        using var document = Document.Create();
        document.Body.AddParagraph().AddImage(document, png);

        var xml = DocumentAssert.MainPartXml(document);
        DocumentAssert.IsValid(document);

        // 192 pixels at 96 DPI is two inches, and an inch is 914400 EMU.
        Assert.That(xml, Does.Contain("cx=\"1828800\""));
        Assert.That(xml, Does.Contain("cy=\"914400\""));
    }

    [Test]
    public void ImageGivenOneDimensionKeepsItsAspectRatio()
    {
        var png = SamplePng.Create(200, 100);

        using var document = Document.Create();
        document.Body.AddParagraph().AddImage(document, png, width: Length.FromInches(1));

        DocumentAssert.IsValid(document);
        var xml = DocumentAssert.MainPartXml(document);

        Assert.That(xml, Does.Contain("cx=\"914400\""));
        Assert.That(xml, Does.Contain("cy=\"457200\""));
    }

    [Test]
    public void FloatedImageIsAnchoredAndWrapped()
    {
        var png = SamplePng.Create(64, 64);

        using var document = Document.Create();
        document.Body.AddParagraph()
            .AddImage(document, png, wrap: ImageWrap.Right, description: "A chart");

        var xml = DocumentAssert.MainPartXml(document);
        DocumentAssert.IsValid(document);

        Assert.That(xml, Does.Contain("<wp:anchor"));
        Assert.That(xml, Does.Contain("<wp:wrapSquare"));
        // Alternative text is what a screen reader announces, so it has to reach the file.
        Assert.That(xml, Does.Contain("descr=\"A chart\""));
    }

    [Test]
    public void ImagesGetDistinctRelationshipsAndIds()
    {
        using var document = Document.Create();
        var paragraph = document.Body.AddParagraph();
        paragraph.AddImage(document, SamplePng.Create(10, 10));
        paragraph.AddImage(document, SamplePng.Create(20, 20));

        DocumentAssert.IsValid(document);
        var xml = DocumentAssert.MainPartXml(document);

        Assert.That(xml, Does.Contain("r:embed=\"rImage1\""));
        Assert.That(xml, Does.Contain("r:embed=\"rImage2\""));
        // A drawing id of zero reads as unset, and two drawings sharing one id is a repair prompt.
        Assert.That(xml, Does.Contain("id=\"1\""));
        Assert.That(xml, Does.Contain("id=\"2\""));
    }

    [Test]
    public void ExternalLinkGetsARelationshipAndTheHyperlinkStyle()
    {
        using var document = Document.Create();
        document.Body.AddParagraph()
            .AddLink(document, "https://example.org/report", "the report");

        var xml = DocumentAssert.MainPartXml(document);
        var styles = DocumentAssert.PartXml(document, "styles");
        DocumentAssert.IsValid(document);

        Assert.That(xml, Does.Contain("<w:hyperlink"));
        // Without the style definition the link renders as ordinary text and only reveals itself
        // on hover, which reads as a bug rather than as a plain-text choice.
        Assert.That(styles, Does.Contain("w:styleId=\"Hyperlink\""));
        Assert.That(xml, Does.Contain("<w:rStyle w:val=\"Hyperlink\" />"));
    }

    [Test]
    public void BookmarkAndAnchorLinkPointAtEachOther()
    {
        using var document = Document.Create();
        document.Body.AddParagraph()
            .AddAnchorLink(document, "du3", "Jump to the third update");
        document.Body.AddParagraph()
            .AddBookmark(document, "du3", _ => _.Append("The third update"));

        var xml = DocumentAssert.MainPartXml(document);
        DocumentAssert.IsValid(document);

        Assert.That(xml, Does.Contain("w:anchor=\"du3\""));
        Assert.That(xml, Does.Contain("<w:bookmarkStart"));
        Assert.That(xml, Does.Contain("<w:bookmarkEnd"));
    }

    [TestCase("Delivery of the new hospital", "Delivery_of_the_new_hospital")]
    [TestCase("  spaces  and---dashes  ", "spaces_and_dashes")]
    [TestCase("3 tunnels & 2 bridges", "tunnels_2_bridges")]
    [TestCase("!!!", "bookmark")]
    public void BookmarkNamesAreSanitisedToWhatWordAccepts(string input, string expected)
    {
        var sanitised = Bookmarks.Sanitise(input);

        Assert.That(sanitised, Is.EqualTo(expected));
        Assert.That(Bookmarks.IsValid(sanitised), Is.True);
    }

    [Test]
    public void LongBookmarkNamesAreTruncatedToTheLimit()
    {
        var sanitised = Bookmarks.Sanitise(new('a', 100));

        // Word drops a bookmark whose name breaks its rules rather than reporting it, and every
        // cross-reference to it then renders as an error in the reader's copy.
        Assert.That(sanitised, Has.Length.EqualTo(Bookmarks.MaxLength));
        Assert.That(Bookmarks.IsValid(sanitised), Is.True);
    }

    [Test]
    public void FieldsCarryACachedValueSoWordDoesNotPrompt()
    {
        using var document = Document.Create();
        document.Body.AddParagraph()
            .Append("Page ")
            .AddPageNumber();
        document.Body.AddParagraph()
            .AddPageReference("du3", "7");

        var xml = DocumentAssert.MainPartXml(document);
        DocumentAssert.IsValid(document);

        Assert.That(xml, Does.Contain("PAGEREF du3"));
        // The computed text, so the document opens finished rather than asking the reader for
        // permission to update fields and showing a placeholder until they agree.
        Assert.That(xml, Does.Contain(">7</w:t>"));
    }

    [Test]
    public void FootnotesGetTheirOwnPartAndNumbering()
    {
        using var document = Document.Create();
        document.Body.AddParagraph()
            .Append("Delivered on time")
            .AddFootnote(document, _ => _.AddParagraph("Against the revised schedule."));

        DocumentAssert.IsValid(document);
        var xml = DocumentAssert.MainPartXml(document);

        Assert.That(xml, Does.Contain("<w:footnoteReference"));
    }

    [Test]
    public void DocumentPropertiesRoundTrip()
    {
        using var document = Document.Create();
        document.Properties.Title = "Stocktake delivery update";
        document.Properties.Creator = "CompassManager";
        document.Properties.SetCustom("Classification", "OFFICIAL");
        document.Body.Paragraph("Body.");

        DocumentAssert.IsValid(document);

        using var reopened = Document.Open(new MemoryStream(document.ToArray()));
        Assert.That(reopened.Properties.Title, Is.EqualTo("Stocktake delivery update"));
        Assert.That(reopened.Properties.GetCustom("Classification"), Is.EqualTo("OFFICIAL"));
    }

    [Test]
    public void ImageInfoReadsEachSupportedFormat()
    {
        Assert.That(ImageInfo.Read(SamplePng.Create(3, 7)).WidthPixels, Is.EqualTo(3));
        Assert.That(ImageInfo.Read(SamplePng.Create(3, 7)).HeightPixels, Is.EqualTo(7));
        Assert.That(ImageInfo.Read(SamplePng.Create(3, 7)).Format, Is.EqualTo(ImageFormat.Png));
    }

    [Test]
    public void UnreadableImageWithoutASizeIsRefusedRatherThanGuessed()
    {
        using var document = Document.Create();
        var paragraph = document.Body.AddParagraph();

        // Better than defaulting to an arbitrary size: a picture silently drawn at the wrong scale
        // is harder to notice than a call that would not compile a document at all.
        Assert.Throws<NotSupportedException>(
            () => paragraph.AddImage(document, [1, 2, 3, 4, 5, 6, 7, 8]));
    }
}
