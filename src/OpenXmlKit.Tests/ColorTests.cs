/// <summary>
/// Colour parsing, and the span overload the HTML-facing consumers need.
/// </summary>
[TestFixture]
public class ColorTests
{
    [Test]
    public void ParsesTheFormsWordAndCssBothUse()
    {
        Assert.That(Color.Parse("#C00000"), Is.EqualTo(Color.FromRgb(0xC00000)));
        Assert.That(Color.Parse("C00000"), Is.EqualTo(Color.FromRgb(0xC00000)));
        // #RGB is shorthand for #RRGGBB, each digit doubled.
        Assert.That(Color.Parse("#0AF"), Is.EqualTo(Color.FromRgb(0x00AAFF)));
        Assert.That(Color.Parse("auto").IsAuto, Is.True);
        Assert.That(Color.Parse("  #C00000  "), Is.EqualTo(Color.FromRgb(0xC00000)));
    }

    [Test]
    public void RejectsWhatItCannotRead()
    {
        Assert.That(Color.TryParse("", out _), Is.False);
        Assert.That(Color.TryParse("#12345", out _), Is.False);
        Assert.That(Color.TryParse("#GGGGGG", out _), Is.False);
        Assert.That(Color.TryParse((string?) null, out _), Is.False);
    }

    [Test]
    public void ASliceParsesWithoutBeingMaterialised()
    {
        // The shape the CSS parsers hand over: a colour is a slice of a declaration, which is a
        // slice of the style attribute it came from. None of those need to become strings.
        const string declaration = "color:#C00000;font-weight:bold";
        var value = declaration.AsSpan("color:".Length, "#C00000".Length);

        Assert.That(Color.TryParse(value, out var color), Is.True);
        Assert.That(color, Is.EqualTo(Color.FromRgb(0xC00000)));
    }

    [Test]
    public void TheTwoOverloadsAgree()
    {
        string[] values = ["#C00000", "C00000", "#0AF", "auto", "  #FFF  ", "", "#12345", "nonsense"];
        foreach (var value in values)
        {
            var fromString = Color.TryParse(value, out var stringColor);
            var fromSpan = Color.TryParse(value.AsSpan(), out var spanColor);

            Assert.That(fromSpan, Is.EqualTo(fromString), value);
            Assert.That(spanColor, Is.EqualTo(stringColor), value);
        }
    }

    [Test]
    public void AnEmptySpanIsNotAColour()
    {
        // The string overload has null to reject; the span overload has only emptiness.
        Assert.That(Color.TryParse(default(ReadOnlySpan<char>), out var color), Is.False);
        Assert.That(color.IsAuto, Is.True);
    }

    [Test]
    public void AParsedColourSurvivesTheRoundTrip()
    {
        using var document = Document.Create();
        document.Body.AddParagraph()
            .AddRun("text").Font.Color = Color.Parse("#C00000");

        var bytes = DocumentAssert.IsValid(document);

        using var read = DocumentView.Open(bytes);
        Assert.That(read.Body.Paragraphs.Single().Runs.Single().Font.Color, Is.EqualTo(Color.FromRgb(0xC00000)));
    }
}
