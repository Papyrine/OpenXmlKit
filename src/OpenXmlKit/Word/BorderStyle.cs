namespace OpenXmlKit.Word;

/// <summary>
/// The line a border draws. Width and colour are stated separately, on <see cref="Border"/>.
/// </summary>
public enum BorderStyle
{
    /// <summary>
    /// Explicitly no border, overriding an inherited one.
    /// </summary>
    None,

    /// <summary>
    /// One solid line, and the border almost every document wants.
    /// </summary>
    Single,

    /// <summary>
    /// One solid line, heavier than <see cref="Single"/> at the same stated width.
    /// </summary>
    Thick,

    /// <summary>
    /// Two parallel lines.
    /// </summary>
    Double,

    /// <summary>
    /// A single dotted line.
    /// </summary>
    Dotted,

    /// <summary>
    /// A single dashed line.
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
    /// Three parallel lines.
    /// </summary>
    Triple,

    /// <summary>
    /// A single wavy line.
    /// </summary>
    Wave,

    /// <summary>
    /// Two parallel wavy lines.
    /// </summary>
    DoubleWave,

    /// <summary>
    /// A bevel that makes the enclosed area look recessed.
    /// </summary>
    Inset,

    /// <summary>
    /// A bevel that makes it look raised.
    /// </summary>
    Outset
}
