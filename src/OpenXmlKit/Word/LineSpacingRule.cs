namespace OpenXmlKit.Word;

/// <summary>
/// How the space between lines in a paragraph is measured.
/// </summary>
public enum LineSpacingRule
{
    /// <summary>
    /// A multiple of single line spacing.
    /// </summary>
    Multiple,

    /// <summary>
    /// At least this much, growing for taller content.
    /// </summary>
    AtLeast,

    /// <summary>
    /// Exactly this much, clipping taller content.
    /// </summary>
    Exactly
}
