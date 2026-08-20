namespace OpenXmlKit.Word;

/// <summary>
/// Raised or lowered text that is also reduced in size. For a raise that keeps the size,
/// use <see cref="Font.Position"/>.
/// </summary>
public enum VerticalTextPosition
{
    /// <summary>
    /// On the baseline, overriding an inherited raise or lower.
    /// </summary>
    Baseline,

    /// <summary>
    /// Raised and shrunk.
    /// </summary>
    Superscript,

    /// <summary>
    /// Lowered and shrunk.
    /// </summary>
    Subscript
}
