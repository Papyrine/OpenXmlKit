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
    public void ArgbHexIsExcelsEightDigitAlphaFirstForm()
    {
        Assert.That(Color.Parse("#C00000").ToArgbHex(), Is.EqualTo("FFC00000"));
        Assert.That(Color.FromRgb(0x00AAFF).ToArgbHex(), Is.EqualTo("FF00AAFF"));

        // Word's own six-digit form is what Value carries; the two differ only by the alpha byte
        // Excel insists on and Word has no place for.
        Assert.That(Color.Parse("#C00000").ToString(), Is.EqualTo("#C00000"));
    }

    [Test]
    public void AColourWithNoRgbHasNoArgb()
    {
        // Assigning null to an Rgb attribute leaves it out, which is what both of these mean.
        Assert.That(Color.Auto.ToArgbHex(), Is.Null);
        Assert.That(Color.FromTheme(ThemeColor.Accent1).ToArgbHex(), Is.Null);
    }

    [Test]
    public void EightDigitInputIsNotParsedBecauseItsOrderIsAmbiguous()
    {
        // Excel writes AARRGGBB and CSS writes RRGGBBAA. Nothing in the string says which, so
        // reading it either way would silently swap a channel for the alpha on half the inputs.
        Assert.That(Color.TryParse("FFC00000", out _), Is.False);
        Assert.That(Color.TryParse("#FFC00000", out _), Is.False);
    }

    [Test]
    public void ArgbIsReadAlphaFirst()
    {
        Assert.That(Color.TryParseArgb("FFC00000", out var opaque), Is.True);
        Assert.That(opaque, Is.EqualTo(Color.FromRgb(0xC00000)));

        Assert.That(Color.TryParseArgb("#FFEFEFEF", out var hashed), Is.True);
        Assert.That(hashed, Is.EqualTo(Color.FromRgb(0xEFEFEF)));

        // Read for validity, then dropped: a Color has no alpha to keep, and ToArgbHex writes FF
        // back regardless.
        Assert.That(Color.TryParseArgb("80FFFFFF", out var translucent), Is.True);
        Assert.That(translucent, Is.EqualTo(Color.White));
        Assert.That(translucent.ToArgbHex(), Is.EqualTo("FFFFFFFF"));
    }

    [Test]
    public void ArgbAlsoReadsEverythingTryParseDoes()
    {
        // The single entry point for a value that may or may not carry an alpha byte, which is
        // what a consumer accepting both forms needs.
        string[] values = ["#C00000", "C00000", "#0AF", "auto", "  #FFF  "];
        foreach (var value in values)
        {
            Assert.That(Color.TryParseArgb(value, out var argb), Is.True, value);
            Color.TryParse(value, out var plain);
            Assert.That(argb, Is.EqualTo(plain), value);
        }
    }

    [Test]
    public void RubbishInTheAlphaPositionIsNotIgnored()
    {
        // Slicing the first two characters off unexamined would let this parse as C00000.
        Assert.That(Color.TryParseArgb("ZZC00000", out _), Is.False);
        Assert.That(Color.TryParseArgb((string?) null, out _), Is.False);
        Assert.That(Color.TryParseArgb("FFC0000", out _), Is.False);
    }

    [Test]
    public void ArgbRoundTripsThroughToArgbHex()
    {
        var source = Color.FromRgb(0xC00000);
        Assert.That(Color.TryParseArgb(source.ToArgbHex(), out var parsed), Is.True);
        Assert.That(parsed, Is.EqualTo(source));
    }

    [Test]
    public void ThePlainParseStillRefusesEightDigits()
    {
        // The two methods differ precisely here, which is the reason there are two.
        Assert.That(Color.TryParse("FFC00000", out _), Is.False);
        Assert.That(Color.TryParseArgb("FFC00000", out _), Is.True);
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
