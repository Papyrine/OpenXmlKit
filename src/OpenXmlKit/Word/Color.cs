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
        color = Auto;
        if (value == null)
        {
            return false;
        }

        var span = value.AsSpan().Trim();
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
