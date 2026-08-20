namespace OpenXmlKit.Word;

/// <summary>
/// How text is laid out inside a table cell.
/// </summary>
public enum TextDirection
{
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
