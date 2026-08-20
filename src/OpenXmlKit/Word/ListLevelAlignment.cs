namespace OpenXmlKit.Word;

/// <summary>
/// How a list marker sits against the position reserved for it.
/// </summary>
public enum ListLevelAlignment
{
    /// <summary>
    /// The marker starts at the indent.
    /// </summary>
    Left,

    /// <summary>
    /// The marker is centred on the indent.
    /// </summary>
    Center,

    /// <summary>
    /// The marker ends at the indent, which keeps roman numerals aligned as they lengthen.
    /// </summary>
    Right
}
