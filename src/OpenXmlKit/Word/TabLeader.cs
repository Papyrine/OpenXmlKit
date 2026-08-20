namespace OpenXmlKit.Word;

/// <summary>
/// What fills the gap a tab jumps across — the row of dots in a table of contents, typically.
/// </summary>
public enum TabLeader
{
    /// <summary>
    /// Nothing between the text and the tab stop.
    /// </summary>
    None,

    /// <summary>
    /// A run of dots, which is what a table of contents uses to reach the page number.
    /// </summary>
    Dots,

    /// <summary>
    /// A run of hyphens.
    /// </summary>
    Dashes,

    /// <summary>
    /// A continuous line along the baseline.
    /// </summary>
    Underscore,

    /// <summary>
    /// A continuous heavy line.
    /// </summary>
    Heavy,

    /// <summary>
    /// A run of dots at mid height rather than on the baseline.
    /// </summary>
    MiddleDot
}
