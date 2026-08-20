namespace OpenXmlKit.Word;

/// <summary>
/// How text behaves around a picture.
/// </summary>
public enum ImageWrap
{
    /// <summary>
    /// The picture sits in the line of text like an oversized character.
    /// </summary>
    Inline,

    /// <summary>
    /// The picture floats at the left margin and text wraps around it.
    /// </summary>
    Left,

    /// <summary>
    /// The picture floats at the right margin and text wraps around it.
    /// </summary>
    Right
}
