namespace OpenXmlKit.Word;

/// <summary>
/// Which pages of a section a header or footer applies to.
/// </summary>
/// <remarks>
/// <see cref="First"/> and <see cref="Even"/> only take effect when the section opts into them —
/// see <c>PageSetup.DifferentFirstPage</c> and <c>PageSetup.DifferentOddAndEvenPages</c>. Adding
/// one without the flag produces a header Word never shows.
/// </remarks>
public enum HeaderFooterKind
{
    /// <summary>
    /// Every page the more specific kinds do not claim.
    /// </summary>
    Default,

    /// <summary>
    /// Used on the first page of the section, and rendered only when the section says so, which
    /// <see cref="Section.AddHeader"/> handles.
    /// </summary>
    First,

    /// <summary>
    /// Used on even pages, and rendered only when the document-wide odd/even switch is on.
    /// </summary>
    Even
}
