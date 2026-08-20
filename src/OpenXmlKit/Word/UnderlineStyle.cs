namespace OpenXmlKit.Word;

public enum UnderlineStyle
{
    /// <summary>
    /// Explicitly not underlined, overriding an inherited underline.
    /// </summary>
    None,

    Single,
    Double,
    Thick,
    Dotted,
    Dashed,
    DotDash,
    DotDotDash,
    Wave,

    /// <summary>
    /// Underlines the words but not the spaces between them.
    /// </summary>
    Words
}
