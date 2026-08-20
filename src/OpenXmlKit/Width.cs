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

    /// <summary>
    /// A fixed width in points.
    /// </summary>
    public static Width FromPoints(double points) =>
        new(WidthUnit.Absolute, points);

    /// <summary>
    /// A fixed width in twips, the unit the file format stores.
    /// </summary>
    public static Width FromTwips(double twips) =>
        new(WidthUnit.Absolute, twips / 20);

    /// <summary>
    /// A fixed width from a <see cref="Length"/>.
    /// </summary>
    public static Width From(Length length) =>
        new(WidthUnit.Absolute, length.TotalPoints);

    /// <summary>
    /// What the width is measured in, which decides how it is written.
    /// </summary>
    public WidthUnit Unit => unit;

    /// <summary>
    /// The value in this width's own unit: a percentage for <see cref="WidthUnit.Percent"/>, points
    /// for <see cref="WidthUnit.Absolute"/>, and zero for <see cref="WidthUnit.Auto"/>.
    /// </summary>
    public double Value => value;

    /// <summary>
    /// Whether the width is left to Word, which is what an unstated width means.
    /// </summary>
    public bool IsAuto => unit == WidthUnit.Auto;

    /// <summary>
    /// The width as a <see cref="Length"/>. Meaningless unless <see cref="Unit"/> is
    /// <see cref="WidthUnit.Absolute"/>.
    /// </summary>
    public Length AsLength =>
        unit == WidthUnit.Absolute ? Length.FromPoints(value) : Length.Zero;

    // Word encodes a percentage as fiftieths of a percent, so 100% is 5000.
    internal int FiftiethsOfAPercent =>
        (int) Math.Round(value * 50, MidpointRounding.AwayFromZero);

    /// <summary>
    /// Whether the two widths state the same thing in the same unit.
    /// </summary>
    public bool Equals(Width other) =>
        unit == other.unit && value.Equals(other.value);

    /// <summary>
    /// Whether the other object is a value of this type and equal to this one.
    /// </summary>
    public override bool Equals(object? obj) =>
        obj is Width other && Equals(other);

    /// <summary>
    /// A hash consistent with equality.
    /// </summary>
    public override int GetHashCode() =>
        (unit, value).GetHashCode();

    /// <summary>
    /// Whether the two widths state the same thing in the same unit.
    /// </summary>
    public static bool operator ==(Width left, Width right) => left.Equals(right);

    /// <summary>
    /// Whether the two widths differ.
    /// </summary>
    public static bool operator !=(Width left, Width right) => !left.Equals(right);

    /// <summary>
    /// A readable form, for logs and debugging rather than for the file.
    /// </summary>
    public override string ToString() =>
        unit switch
        {
            WidthUnit.Percent => value.ToString("0.##", CultureInfo.InvariantCulture) + "%",
            WidthUnit.Absolute => value.ToString("0.##", CultureInfo.InvariantCulture) + "pt",
            _ => "auto"
        };
}
