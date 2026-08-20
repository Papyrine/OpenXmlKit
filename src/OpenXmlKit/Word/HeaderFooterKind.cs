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

    First,
    Even
}
