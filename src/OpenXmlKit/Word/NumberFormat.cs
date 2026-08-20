namespace OpenXmlKit.Word;

/// <summary>
/// How a list level labels its items.
/// </summary>
public enum NumberFormat
{
    /// <summary>
    /// No marker at all.
    /// </summary>
    None,

    /// <summary>
    /// A glyph rather than a number — the level's text supplies which one.
    /// </summary>
    Bullet,

    /// <summary>
    /// 1, 2, 3.
    /// </summary>
    Decimal,

    /// <summary>
    /// I, II, III.
    /// </summary>
    UpperRoman,

    /// <summary>
    /// i, ii, iii.
    /// </summary>
    LowerRoman,

    /// <summary>
    /// A, B, C.
    /// </summary>
    UpperLetter,

    /// <summary>
    /// a, b, c.
    /// </summary>
    LowerLetter,

    /// <summary>
    /// 1st, 2nd, 3rd.
    /// </summary>
    Ordinal,

    /// <summary>
    /// One, Two, Three.
    /// </summary>
    CardinalText,

    /// <summary>
    /// First, Second, Third.
    /// </summary>
    OrdinalText,

    /// <summary>
    /// Decimal with a leading zero below ten.
    /// </summary>
    DecimalZero
}
