namespace OpenXmlKit.Word;

/// <summary>
/// How a paragraph sits between its margins.
/// </summary>
public enum ParagraphAlignment
{
    /// <summary>
    /// Flush left, ragged right.
    /// </summary>
    Left,

    /// <summary>
    /// Centred, ragged on both sides.
    /// </summary>
    Center,

    /// <summary>
    /// Flush right, ragged left.
    /// </summary>
    Right,

    /// <summary>
    /// Flush to both margins, spacing stretched to reach. Word calls this <c>both</c>.
    /// </summary>
    Justify,

    /// <summary>
    /// Justified, and the last line stretched as well rather than left short.
    /// </summary>
    Distribute
}
