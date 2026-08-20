namespace OpenXmlKit.Word;

/// <summary>
/// The Word styles this library can supply a stock definition for.
/// </summary>
/// <remarks>
/// Word carries these at application level and only writes them into a document's styles part when
/// the user inserts something that uses them. A document built in code therefore names a style that
/// is not there, and renders unstyled — which is why two libraries in this estate each grew their
/// own copy of the same seeding code. <see cref="Styles.EnsureBuiltIn(BuiltInStyle)"/> is that code, once.
/// <para>
/// Referring to a style by this enum rather than by its id also survives localisation: Word's
/// display names are translated, the ids are not.
/// </para>
/// </remarks>
public enum BuiltInStyle
{
    Normal,
    Heading1,
    Heading2,
    Heading3,
    Heading4,
    Heading5,
    Heading6,
    Title,
    Subtitle,
    Quote,
    IntenseQuote,
    Caption,
    Header,
    Footer,
    ListParagraph,
    NoSpacing,
    Hyperlink,
    FootnoteText,
    FootnoteReference,

    /// <summary>
    /// The default table style, which supplies cell padding and nothing else.
    /// </summary>
    TableNormal,

    /// <summary>
    /// Single-line borders, inheriting cell padding from <see cref="TableNormal"/>. What Word
    /// applies to a table inserted from the ribbon.
    /// </summary>
    TableGrid
}
