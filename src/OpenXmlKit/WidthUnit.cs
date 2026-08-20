namespace OpenXmlKit;

/// <summary>
/// What a <see cref="Width"/> is measured in.
/// </summary>
public enum WidthUnit
{
    /// <summary>
    /// Sized by content.
    /// </summary>
    Auto,

    /// <summary>
    /// A share of the containing width.
    /// </summary>
    Percent,

    /// <summary>
    /// A fixed measurement.
    /// </summary>
    Absolute
}
