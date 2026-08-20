namespace OpenXmlKit.Word;

/// <summary>
/// What separates a list marker from the text beside it.
/// </summary>
public enum ListTrailingCharacter
{
    /// <summary>
    /// A tab, which lines the text up on the level indent.
    /// </summary>
    Tab,
    /// <summary>
    /// A single space.
    /// </summary>
    Space,
    /// <summary>
    /// Nothing at all, so the text begins immediately after the marker.
    /// </summary>
    Nothing
}
