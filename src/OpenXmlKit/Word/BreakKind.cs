namespace OpenXmlKit.Word;

/// <summary>
/// What a break inside a run breaks.
/// </summary>
public enum BreakKind
{
    /// <summary>
    /// A new line within the same paragraph.
    /// </summary>
    Line,

    /// <summary>
    /// Moves what follows onto the next page.
    /// </summary>
    Page,
    /// <summary>
    /// Moves what follows into the next column.
    /// </summary>
    Column
}
