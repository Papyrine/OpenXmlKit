namespace OpenXmlKit;

/// <summary>
/// A table or cell width: a number and the thing that number is a number of.
/// </summary>
/// <remarks>
/// <c>w:tblW</c> and <c>w:tcW</c> carry a value and a type together, and the value means nothing
/// without it — 5000 is 100% under <c>pct</c> and three and a half inches under <c>dxa</c>. Pairing
/// them in one type is what stops <c>Width = "5000"</c> from being a coin toss, and is why the
/// fiftieths-of-a-percent encoding never has to appear at a call site.
/// </remarks>
public readonly struct Width :
    IEquatable<Width>
{
    readonly WidthUnit unit;
    readonly double value;

    Width(WidthUnit unit, double value)
    {
        this.unit = unit;
        this.value = value;
    }

    /// <summary>
    /// Sized by content. The default.
    /// </summary>
    public static Width Auto => default;

    /// <summary>
    /// A share of the containing width, as a percentage — <c>Width.Percent(22)</c> for
    /// <c>width:22%</c>.
    /// </summary>
    public static Width Percent(double percent) =>
        new(WidthUnit.Percent, percent);

    public static Width FromPoints(double points) =>
        new(WidthUnit.Absolute, points);

    public static Width FromTwips(double twips) =>
        new(WidthUnit.Absolute, twips / 20);

    public static Width From(Length length) =>
        new(WidthUnit.Absolute, length.TotalPoints);

    public WidthUnit Unit => unit;

    /// <summary>
    /// The value in this width's own unit: a percentage for <see cref="WidthUnit.Percent"/>, points
    /// for <see cref="WidthUnit.Absolute"/>, and zero for <see cref="WidthUnit.Auto"/>.
    /// </summary>
    public double Value => value;

    public bool IsAuto => unit == WidthUnit.Auto;

    public Length AsLength =>
        unit == WidthUnit.Absolute ? Length.FromPoints(value) : Length.Zero;

    // Word encodes a percentage as fiftieths of a percent, so 100% is 5000.
    internal int FiftiethsOfAPercent =>
        (int) Math.Round(value * 50, MidpointRounding.AwayFromZero);

    public bool Equals(Width other) =>
        unit == other.unit && value.Equals(other.value);

    public override bool Equals(object? obj) =>
        obj is Width other && Equals(other);

    public override int GetHashCode() =>
        (unit, value).GetHashCode();

    public static bool operator ==(Width left, Width right) => left.Equals(right);
    public static bool operator !=(Width left, Width right) => !left.Equals(right);

    public override string ToString() =>
        unit switch
        {
            WidthUnit.Percent => value.ToString("0.##", CultureInfo.InvariantCulture) + "%",
            WidthUnit.Absolute => value.ToString("0.##", CultureInfo.InvariantCulture) + "pt",
            _ => "auto"
        };
}
