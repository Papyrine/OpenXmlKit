namespace OpenXmlKit.Word;

/// <summary>
/// A colour: an explicit RGB value, a reference to a slot in the document theme, or
/// <see cref="Auto"/> — which lets Word pick, and is what an unset colour means.
/// </summary>
public readonly struct Color :
    IEquatable<Color>
{
    readonly int rgb;
    readonly ThemeColor theme;
    readonly bool isSet;

    Color(int rgb, ThemeColor theme, bool isSet)
    {
        this.rgb = rgb;
        this.theme = theme;
        this.isSet = isSet;
    }

    /// <summary>
    /// Word's own choice — usually black on a light background, and white on a dark one. The
    /// default, and what an unstated colour resolves to.
    /// </summary>
    public static Color Auto => default;

    public static Color FromRgb(byte red, byte green, byte blue) =>
        new((red << 16) | (green << 8) | blue, ThemeColor.None, true);

    /// <summary>
    /// A packed <c>0xRRGGBB</c>. Any alpha byte is ignored — Word run and cell colours are opaque.
    /// </summary>
    public static Color FromRgb(int rgb) =>
        new(rgb & 0xFFFFFF, ThemeColor.None, true);

    /// <summary>
    /// Parses <c>#RRGGBB</c>, <c>RRGGBB</c>, <c>#RGB</c> or the literal <c>auto</c>.
    /// </summary>
    public static Color Parse(string value)
    {
        if (!TryParse(value, out var color))
        {
            throw new FormatException($"Could not read '{value}' as a colour. Expected #RRGGBB, RRGGBB, #RGB or auto.");
        }

        return color;
    }

    public static bool TryParse([NotNullWhen(true)] string? value, out Color color)
    {
        if (value == null)
        {
            color = Auto;
            return false;
        }

        return TryParse(value.AsSpan(), out color);
    }

    /// <summary>
    /// Parses <c>#RRGGBB</c>, <c>RRGGBB</c>, <c>#RGB</c> or the literal <c>auto</c>.
    /// </summary>
    /// <remarks>
    /// The span overload is the implementation, and exists because the parsers that will feed it
    /// already work in spans — a CSS declaration is a slice of the style attribute it came from,
    /// and a colour is a slice of that. Offering only the string form would put a ToString on
    /// every one of those boundaries, for text that is being read rather than kept.
    /// </remarks>
    public static bool TryParse(ReadOnlySpan<char> value, out Color color)
    {
        color = Auto;

        var span = value.Trim();
        if (span.Length == 0)
        {
            return false;
        }

        if (span.Equals("auto".AsSpan(), StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (span[0] == '#')
        {
            span = span[1..];
        }

        if (span.Length == 3)
        {
            // #RGB is shorthand for #RRGGBB, each digit doubled.
            if (!TryHex(span[0], out var r) ||
                !TryHex(span[1], out var g) ||
                !TryHex(span[2], out var b))
            {
                return false;
            }

            color = FromRgb((byte) (r * 17), (byte) (g * 17), (byte) (b * 17));
            return true;
        }

        if (span.Length != 6)
        {
            return false;
        }

        var packed = 0;
        foreach (var character in span)
        {
            if (!TryHex(character, out var digit))
            {
                return false;
            }

            packed = (packed << 4) | digit;
        }

        color = FromRgb(packed);
        return true;
    }

    /// <summary>
    /// Parses Excel's <c>AARRGGBB</c>, and everything <see cref="TryParse(string?, out Color)"/>
    /// reads as well.
    /// </summary>
    /// <remarks>
    /// The separate name is the whole point. Eight hex digits are alpha-first in Excel and
    /// alpha-last in CSS, and nothing in the string says which, so the caller states the order by
    /// choosing the method rather than the parser guessing. Three and six digits mean the same
    /// thing to both and are read here too, so this is the single entry point for a value that may
    /// or may not carry an alpha byte.
    /// <para>
    /// The alpha is read for validity and then dropped, because a <see cref="Color"/> has none to
    /// keep: Word's colours are opaque, and <see cref="ToArgbHex"/> writes <c>FF</c> back. A value
    /// that was half transparent going in comes out solid.
    /// </para>
    /// </remarks>
    public static bool TryParseArgb([NotNullWhen(true)] string? value, out Color color)
    {
        if (value == null)
        {
            color = Auto;
            return false;
        }

        return TryParseArgb(value.AsSpan(), out color);
    }

    /// <inheritdoc cref="TryParseArgb(string?, out Color)"/>
    public static bool TryParseArgb(ReadOnlySpan<char> value, out Color color)
    {
        var span = value.Trim();
        if (span.Length > 0 &&
            span[0] == '#')
        {
            span = span[1..];
        }

        if (span.Length != 8)
        {
            return TryParse(span, out color);
        }

        // The alpha digits are checked rather than skipped: slicing them off unexamined would let
        // an eight-character string with two pieces of rubbish on the front parse as a colour.
        if (!TryHex(span[0], out _) ||
            !TryHex(span[1], out _))
        {
            color = Auto;
            return false;
        }

        return TryParse(span[2..], out color);
    }

    static bool TryHex(char character, out int value)
    {
        if (character is >= '0' and <= '9')
        {
            value = character - '0';
            return true;
        }

        if (character is >= 'a' and <= 'f')
        {
            value = character - 'a' + 10;
            return true;
        }

        if (character is >= 'A' and <= 'F')
        {
            value = character - 'A' + 10;
            return true;
        }

        value = 0;
        return false;
    }

    /// <summary>
    /// A slot in the document theme, so the colour follows the template rather than being pinned.
    /// </summary>
    public static Color FromTheme(ThemeColor theme) =>
        new(0, theme, true);

    public bool IsAuto => !isSet;
    public bool IsTheme => theme != ThemeColor.None;
    public ThemeColor Theme => theme;

    // R/G/B rather than Red/Green/Blue, so the channel accessors do not collide with the named
    // colours below. System.Drawing.Color draws the same line for the same reason.
    public byte R => (byte) ((rgb >> 16) & 0xFF);
    public byte G => (byte) ((rgb >> 8) & 0xFF);
    public byte B => (byte) (rgb & 0xFF);

    /// <summary>
    /// The <c>RRGGBB</c> form Word writes, or <c>auto</c>.
    /// </summary>
    internal string Value =>
        isSet ? rgb.ToString("X6", CultureInfo.InvariantCulture) : "auto";

    /// <summary>
    /// The <c>AARRGGBB</c> form Excel writes, or null for a colour that states no explicit RGB.
    /// </summary>
    /// <remarks>
    /// Word writes six hex digits and treats colours as opaque; Excel writes eight, and puts the
    /// alpha byte first. Null comes back for <see cref="Auto"/> and for a theme reference, because
    /// neither has an RGB value to state — assigning null to an <c>Rgb</c> attribute leaves the
    /// attribute out, which is what both of them mean.
    /// <para>
    /// The ordering is the trap: Excel's eight digits are alpha-first, where CSS's eight-digit
    /// form is alpha-last. The two cannot be told apart by inspection, which is why this is a
    /// one-way conversion and <see cref="TryParse(string?, out Color)"/> reads three or six digits
    /// but never eight.
    /// </para>
    /// </remarks>
    public string? ToArgbHex() =>
        isSet && !IsTheme
            ? "FF" + rgb.ToString("X6", CultureInfo.InvariantCulture)
            : null;

    public static implicit operator Color(string value) => Parse(value);

    public static readonly Color Black = FromRgb(0x000000);
    public static readonly Color White = FromRgb(0xFFFFFF);
    public static readonly Color Red = FromRgb(0xFF0000);
    public static readonly Color Green = FromRgb(0x008000);
    public static readonly Color Blue = FromRgb(0x0000FF);
    public static readonly Color Yellow = FromRgb(0xFFFF00);
    public static readonly Color Gray = FromRgb(0x808080);
    public static readonly Color LightGray = FromRgb(0xD3D3D3);
    public static readonly Color DarkGray = FromRgb(0xA9A9A9);

    public bool Equals(Color other) =>
        isSet == other.isSet &&
        rgb == other.rgb &&
        theme == other.theme;

    public override bool Equals(object? obj) =>
        obj is Color other && Equals(other);

    public override int GetHashCode() =>
        (isSet, rgb, theme).GetHashCode();

    public static bool operator ==(Color left, Color right) => left.Equals(right);
    public static bool operator !=(Color left, Color right) => !left.Equals(right);

    public override string ToString() =>
        IsTheme ? theme.ToString() : isSet ? "#" + Value : "auto";
}
