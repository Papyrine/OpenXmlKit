namespace OpenXmlKit.Word;

/// <summary>
/// Which way round the page is.
/// </summary>
public enum PageOrientation
{
    /// <summary>
    /// Taller than wide.
    /// </summary>
    Portrait,

    /// <summary>
    /// Wider than tall. Setting this swaps the page dimensions rather than only recording a flag.
    /// </summary>
    Landscape
}
