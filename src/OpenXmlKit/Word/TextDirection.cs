namespace OpenXmlKit.Word;

/// <summary>
/// How text is laid out inside a table cell.
/// </summary>
public enum TextDirection
{
    /// <summary>
    /// Ordinary left-to-right text.
    /// </summary>
    Horizontal,

    /// <summary>
    /// Rotated 90° clockwise, reading top to bottom.
    /// </summary>
    RotateDown,

    /// <summary>
    /// Rotated 90° anticlockwise, reading bottom to top.
    /// </summary>
    RotateUp
}
