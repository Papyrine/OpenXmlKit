namespace OpenXmlKit.Word;

/// <summary>
/// Where a section begins relative to the one before it.
/// </summary>
public enum SectionStart
{
    /// <summary>
    /// Carries on down the same page, which is how a section changes column count mid-page.
    /// </summary>
    Continuous,

    /// <summary>
    /// Starts on the next page.
    /// </summary>
    NewPage,

    /// <summary>
    /// Starts at the top of the next column.
    /// </summary>
    NewColumn,

    /// <summary>
    /// Starts on the next even page, leaving a blank one if need be.
    /// </summary>
    EvenPage,

    /// <summary>
    /// Starts on the next odd page, which is how a chapter always opens on a right-hand page.
    /// </summary>
    OddPage
}
