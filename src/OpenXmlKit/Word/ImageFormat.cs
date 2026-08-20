namespace OpenXmlKit.Word;

/// <summary>
/// The image encodings a picture part can hold.
/// </summary>
public enum ImageFormat
{
    /// <summary>
    /// Lossless, with transparency. What this library assumes when a header cannot be read.
    /// </summary>
    Png,

    /// <summary>
    /// Lossy, no transparency.
    /// </summary>
    Jpeg,

    /// <summary>
    /// Indexed colour, with transparency.
    /// </summary>
    Gif,

    /// <summary>
    /// Uncompressed.
    /// </summary>
    Bmp,

    /// <summary>
    /// Read by Word, but not measurable here, so a picture in this format needs an explicit size.
    /// </summary>
    Tiff
}
