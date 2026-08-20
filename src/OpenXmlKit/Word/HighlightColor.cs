namespace OpenXmlKit.Word;

/// <summary>
/// The fixed palette Word's text highlighter offers. Unlike shading, a highlight cannot be an
/// arbitrary colour.
/// </summary>
public enum HighlightColor
{
    /// <summary>
    /// No highlight, overriding an inherited one.
    /// </summary>
    None,

    /// <summary>
    /// Black, which needs light text over it to stay readable.
    /// </summary>
    Black,

    /// <summary>
    /// Word's blue highlighter.
    /// </summary>
    Blue,

    /// <summary>
    /// Word's cyan highlighter.
    /// </summary>
    Cyan,

    /// <summary>
    /// Word's green highlighter.
    /// </summary>
    Green,

    /// <summary>
    /// Word's magenta highlighter.
    /// </summary>
    Magenta,

    /// <summary>
    /// Word's red highlighter.
    /// </summary>
    Red,

    /// <summary>
    /// Word's yellow highlighter.
    /// </summary>
    Yellow,

    /// <summary>
    /// White, which reads as no highlight on a white page but still overrides an inherited one.
    /// </summary>
    White,

    /// <summary>
    /// Word's dark blue highlighter.
    /// </summary>
    DarkBlue,

    /// <summary>
    /// Teal.
    /// </summary>
    DarkCyan,

    /// <summary>
    /// Word's dark green highlighter.
    /// </summary>
    DarkGreen,

    /// <summary>
    /// Purple.
    /// </summary>
    DarkMagenta,

    /// <summary>
    /// Word's dark red highlighter.
    /// </summary>
    DarkRed,

    /// <summary>
    /// Olive, which is what Word draws for this rather than a dark yellow.
    /// </summary>
    DarkYellow,

    /// <summary>
    /// Word's dark gray highlighter.
    /// </summary>
    DarkGray,

    /// <summary>
    /// Word's light gray highlighter.
    /// </summary>
    LightGray
}
