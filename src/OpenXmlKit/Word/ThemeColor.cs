namespace OpenXmlKit.Word;

/// <summary>
/// A slot in the document theme's colour scheme.
/// </summary>
public enum ThemeColor
{
    /// <summary>
    /// Not a theme colour.
    /// </summary>
    None,

    /// <summary>
    /// The lighter background slot — white in the default theme.
    /// </summary>
    Background1,

    /// <summary>
    /// The darker text slot — black in the default theme.
    /// </summary>
    Text1,

    /// <summary>
    /// The darker background slot, for banding and fills.
    /// </summary>
    Background2,

    /// <summary>
    /// The lighter text slot, for text over <see cref="Background2"/>.
    /// </summary>
    Text2,

    /// <summary>
    /// The first accent, which is what Word reaches for first when colouring a chart or table.
    /// </summary>
    Accent1,

    /// <summary>
    /// The second accent.
    /// </summary>
    Accent2,

    /// <summary>
    /// The third accent.
    /// </summary>
    Accent3,

    /// <summary>
    /// The fourth accent.
    /// </summary>
    Accent4,

    /// <summary>
    /// The fifth accent.
    /// </summary>
    Accent5,

    /// <summary>
    /// The sixth and last accent.
    /// </summary>
    Accent6,

    /// <summary>
    /// The colour of a link that has not been followed.
    /// </summary>
    Hyperlink,

    /// <summary>
    /// The colour of a link that has.
    /// </summary>
    FollowedHyperlink
}
