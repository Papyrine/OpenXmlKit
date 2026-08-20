namespace OpenXmlKit.Word;

/// <summary>
/// How text is underlined. The colour is stated separately, on <see cref="Font.UnderlineColor"/>.
/// </summary>
public enum UnderlineStyle
{
    /// <summary>
    /// Explicitly not underlined, overriding an inherited underline.
    /// </summary>
    None,

    /// <summary>
    /// One line under the text.
    /// </summary>
    Single,

    /// <summary>
    /// Two lines.
    /// </summary>
    Double,

    /// <summary>
    /// One heavy line.
    /// </summary>
    Thick,

    /// <summary>
    /// A dotted line.
    /// </summary>
    Dotted,

    /// <summary>
    /// A dashed line.
    /// </summary>
    Dashed,

    /// <summary>
    /// Alternating dots and dashes.
    /// </summary>
    DotDash,

    /// <summary>
    /// Two dots then a dash, repeating.
    /// </summary>
    DotDotDash,

    /// <summary>
    /// A wavy line, which is also what Word draws under a spelling error.
    /// </summary>
    Wave,

    /// <summary>
    /// Underlines the words but not the spaces between them.
    /// </summary>
    Words
}
