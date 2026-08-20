namespace OpenXmlKit.Word;

public enum TabAlignment
{
    Left,
    Center,
    Right,

    /// <summary>
    /// Aligns numbers on their decimal point.
    /// </summary>
    Decimal,

    /// <summary>
    /// Draws a vertical bar rather than positioning text.
    /// </summary>
    Bar,

    /// <summary>
    /// Removes a tab stop inherited from the style.
    /// </summary>
    Clear
}
