/// <summary>
/// The characters XML forbids, which nothing rejects until Save.
/// </summary>
[TestFixture]
public class XmlCharsTests
{
    [Test]
    public void CleanTextIsReturnedUntouched()
    {
        const string value = "Ordinary text\twith a tab\nand a newline.";
        Assert.That(XmlChars.Strip(value), Is.SameAs(value));
        Assert.That(XmlChars.IsValid(value), Is.True);
    }

    [Test]
    public void ForbiddenControlsAreRemoved()
    {
        Assert.That(XmlChars.Strip("a\u0001b\u001Fc"), Is.EqualTo("abc"));
        // Tab, line feed and carriage return are the three controls XML keeps.
        Assert.That(XmlChars.Strip("a\tb\nc\rd"), Is.EqualTo("a\tb\nc\rd"));
    }

    [Test]
    public void APairedSurrogateSurvivesAndALoneOneDoesNot()
    {
        // A matched pair is one character and has to be stepped over together; testing either half
        // alone would drop a perfectly valid emoji.
        Assert.That(XmlChars.Strip("a\U0001F600b"), Is.EqualTo("a\U0001F600b"));
        Assert.That(XmlChars.Strip("a\uD83Db"), Is.EqualTo("ab"));
        Assert.That(XmlChars.Strip("a\uDE00b"), Is.EqualTo("ab"));
    }

    [Test]
    public void NonCharactersAreRemoved()
    {
        Assert.That(XmlChars.Strip("a\uFFFE\uFFFFb"), Is.EqualTo("ab"));
        Assert.That(XmlChars.Strip("a\uFFFDb"), Is.EqualTo("a\uFFFDb"));
    }

    [Test]
    public void CleanSpansComeBackAsTheSameMemory()
    {
        // The contract the span overload exists for: a parser holding a slice of its input pays
        // nothing for text that needs no changing.
        var source = "Ordinary text".AsSpan();
        Assert.That(XmlChars.Strip(source) == source, Is.True);

        var dirty = "a\u0001b".AsSpan();
        Assert.That(XmlChars.Strip(dirty) == dirty, Is.False);
        Assert.That(XmlChars.Strip(dirty).ToString(), Is.EqualTo("ab"));
    }

    [Test]
    public void ASurrogatePairDoesNotForceACopy()
    {
        // The scan cannot tell a paired surrogate from a lone one, so it stops on both; only the
        // walk can, and having decided nothing is wrong it must hand back the original memory.
        var source = "a\U0001F600b".AsSpan();
        Assert.That(XmlChars.Strip(source) == source, Is.True);
    }

    [Test]
    public void TheTwoOverloadsAgree()
    {
        string[] values = ["", "plain", "a\u0001b", "a\uD83Db", "a\U0001F600b", "\uFFFE", "\t\n\r"];
        foreach (var value in values)
        {
            Assert.That(XmlChars.Strip(value.AsSpan()).ToString(), Is.EqualTo(XmlChars.Strip(value)), value);
        }
    }

    [Test]
    public void TextWrittenThroughTheBuildApiIsStripped()
    {
        using var document = Document.Create();
        document.Body.Paragraph("before\u0001after");
        document.Body.AddParagraph().AddField("PAGE", "12\u0002");

        // Without stripping this throws at save, from somewhere that cannot say which string was
        // at fault.
        var bytes = DocumentAssert.IsValid(document);

        using var read = DocumentView.Open(bytes);
        Assert.That(read.Body.Paragraphs.First().Text, Is.EqualTo("beforeafter"));
        Assert.That(read.Body.Paragraphs.ElementAt(1).Fields.Single().Value, Is.EqualTo("12"));
    }
}
