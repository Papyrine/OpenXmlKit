namespace OpenXmlKit.Word;

/// <summary>
/// How a stated row height is treated.
/// </summary>
public enum RowHeightRule
{
    /// <summary>
    /// Sized by content.
    /// </summary>
    Auto,

    /// <summary>
    /// At least this tall, growing for taller content.
    /// </summary>
    AtLeast,

    /// <summary>
    /// Exactly this tall, clipping taller content.
    /// </summary>
    Exactly
}
