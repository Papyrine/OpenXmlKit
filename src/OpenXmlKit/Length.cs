namespace OpenXmlKit;

/// <summary>
/// A distance, expressed in whatever unit is convenient at the call site and converted to whatever
/// unit the OpenXML attribute being written happens to use.
/// </summary>
/// <remarks>
/// Word measures in five different scales — twips for page and indent geometry, half-points for
/// font size, eighths of a point for border widths, EMUs for drawings, and raw points in a few
/// places — and the SDK types those attributes variously as <c>string</c>, <c>int</c>, <c>uint</c>
/// and <see cref="StringValue"/>. Carrying a distance as a bare number therefore means the unit
/// lives in the reader's head, and the classic result is a border eight times too thick or a font
/// at half its intended size.
/// <para>
/// Stored as points because no integer unit divides all five: an eighth of a point is 1587.5 EMU,
/// so an EMU-based integer would lose border widths. Points as a double round-trips every scale
/// the format uses.
/// </para>
/// </remarks>
public readonly struct Length :
    IEquatable<Length>,
    IComparable<Length>
{
    readonly double points;

    Length(double points) =>
        this.points = points;

    /// <summary>
    /// A length in points, the unit this type stores.
    /// </summary>
    public static Length FromPoints(double points) =>
        new(points);

    /// <summary>
    /// Half-points, the unit <c>w:sz</c> uses for font size — 24 half-points is 12pt.
    /// </summary>
    public static Length FromHalfPoints(double halfPoints) =>
        new(halfPoints / 2);

    /// <summary>
    /// Eighths of a point, the unit <c>w:sz</c> uses on a border — 4 eighths is a half-point hairline.
    /// </summary>
    /// <remarks>
    /// The same attribute name as font size, a different unit. Word's own doing, not the SDK's.
    /// </remarks>
    public static Length FromEighthPoints(double eighthPoints) =>
        new(eighthPoints / 8);

    /// <summary>
    /// Twentieths of a point, the unit most page and paragraph geometry uses — 1440 twips is an inch.
    /// </summary>
    public static Length FromTwips(double twips) =>
        new(twips / 20);

    /// <summary>
    /// A length in inches. One inch is 72 points.
    /// </summary>
    public static Length FromInches(double inches) =>
        new(inches * pointsPerInch);

    /// <summary>
    /// A length in centimetres.
    /// </summary>
    public static Length FromCentimeters(double centimeters) =>
        new(centimeters * pointsPerInch / 2.54);

    /// <summary>
    /// A length in millimetres.
    /// </summary>
    public static Length FromMillimeters(double millimeters) =>
        new(millimeters * pointsPerInch / 25.4);

    /// <summary>
    /// Pixels at a given resolution, defaulting to the 96 DPI that html sizes are quoted in.
    /// </summary>
    public static Length FromPixels(double pixels, double dpi = 96) =>
        new(pixels * pointsPerInch / dpi);

    /// <summary>
    /// English Metric Units, the unit DrawingML measures images in — 914400 to the inch.
    /// </summary>
    public static Length FromEmu(double emu) =>
        new(emu / emuPerPoint);

    /// <summary>
    /// No length at all.
    /// </summary>
    public static readonly Length Zero = new(0);

    /// <summary>
    /// The length in points.
    /// </summary>
    public double TotalPoints => points;

    /// <summary>
    /// The length in inches.
    /// </summary>
    public double TotalInches => points / pointsPerInch;

    /// <summary>
    /// The length in centimetres.
    /// </summary>
    public double TotalCentimeters => points * 2.54 / pointsPerInch;

    /// <summary>
    /// The length in millimetres.
    /// </summary>
    public double TotalMillimeters => points * 25.4 / pointsPerInch;

    // The emit-side accessors. Integer, because every attribute they feed is integer-valued, and
    // rounding once here is what stops a half-twip drifting through three conversions.
    internal int Twips => Round(points * 20);
    internal int HalfPoints => Round(points * 2);
    internal int EighthPoints => Round(points * 8);
    internal long Emu => (long) Math.Round(points * emuPerPoint, MidpointRounding.AwayFromZero);

    static int Round(double value) =>
        (int) Math.Round(value, MidpointRounding.AwayFromZero);

    const double pointsPerInch = 72;
    const double emuPerPoint = 12700;

    /// <summary>
    /// Reads a bare number as points, so <c>Size = 13</c> needs no ceremony for the common case.
    /// </summary>
    public static implicit operator Length(double points) =>
        new(points);

    /// <summary>
    /// Treats a bare number as a length in points, so <c>Size = 13</c> reads naturally.
    /// </summary>
    public static implicit operator Length(int points) =>
        new(points);

    /// <summary>
    /// Whether the two lengths are the same.
    /// </summary>
    public bool Equals(Length other) =>
        points.Equals(other.points);

    /// <summary>
    /// Whether the other object is a value of this type and equal to this one.
    /// </summary>
    public override bool Equals(object? obj) =>
        obj is Length other && Equals(other);

    /// <summary>
    /// A hash consistent with equality.
    /// </summary>
    public override int GetHashCode() =>
        points.GetHashCode();

    /// <summary>
    /// Orders by length, so a set of them can be sorted.
    /// </summary>
    public int CompareTo(Length other) =>
        points.CompareTo(other.points);

    /// <summary>
    /// Whether the two lengths are the same.
    /// </summary>
    public static bool operator ==(Length left, Length right) => left.Equals(right);

    /// <summary>
    /// Whether the two lengths differ.
    /// </summary>
    public static bool operator !=(Length left, Length right) => !left.Equals(right);

    /// <summary>
    /// Whether the first length is the shorter.
    /// </summary>
    public static bool operator <(Length left, Length right) => left.points < right.points;

    /// <summary>
    /// Whether the first length is the longer.
    /// </summary>
    public static bool operator >(Length left, Length right) => left.points > right.points;

    /// <summary>
    /// Whether the first length is no longer than the second.
    /// </summary>
    public static bool operator <=(Length left, Length right) => left.points <= right.points;

    /// <summary>
    /// Whether the first length is no shorter than the second.
    /// </summary>
    public static bool operator >=(Length left, Length right) => left.points >= right.points;

    /// <summary>
    /// The sum of two lengths.
    /// </summary>
    public static Length operator +(Length left, Length right) => new(left.points + right.points);

    /// <summary>
    /// The difference between two lengths.
    /// </summary>
    public static Length operator -(Length left, Length right) => new(left.points - right.points);

    /// <summary>
    /// The length scaled by a factor.
    /// </summary>
    public static Length operator *(Length length, double factor) => new(length.points * factor);

    /// <summary>
    /// The length divided by a factor.
    /// </summary>
    public static Length operator /(Length length, double divisor) => new(length.points / divisor);

    /// <summary>
    /// A readable form, for logs and debugging rather than for the file.
    /// </summary>
    public override string ToString() =>
        points.ToString("0.##", CultureInfo.InvariantCulture) + "pt";
}
