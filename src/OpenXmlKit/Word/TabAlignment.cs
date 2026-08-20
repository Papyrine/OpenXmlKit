namespace OpenXmlKit.Word;

/// <summary>
/// How text lines up against a tab stop.
/// </summary>
public enum TabAlignment
{
    /// <summary>
    /// Text begins at the stop.
    /// </summary>
    Left,

    /// <summary>
    /// Text is centred on the stop.
    /// </summary>
    Center,

    /// <summary>
    /// Text ends at the stop.
    /// </summary>
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
