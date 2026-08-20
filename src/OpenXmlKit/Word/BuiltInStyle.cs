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
    /// <summary>
    /// The document default, which every other paragraph style ultimately inherits from.
    /// </summary>
    Normal,

    /// <summary>
    /// The top level of the outline, and what a table of contents picks up first.
    /// </summary>
    Heading1,

    /// <summary>
    /// The second outline level.
    /// </summary>
    Heading2,

    /// <summary>
    /// The third outline level.
    /// </summary>
    Heading3,

    /// <summary>
    /// The fourth outline level.
    /// </summary>
    Heading4,

    /// <summary>
    /// The fifth outline level.
    /// </summary>
    Heading5,

    /// <summary>
    /// The sixth outline level, and the deepest Word offers as a built-in.
    /// </summary>
    Heading6,

    /// <summary>
    /// The document title, which sits outside the heading outline rather than above it.
    /// </summary>
    Title,

    /// <summary>
    /// A line beneath <see cref="Title"/>, in the same spirit.
    /// </summary>
    Subtitle,

    /// <summary>
    /// An indented pull quote.
    /// </summary>
    Quote,

    /// <summary>
    /// A pull quote with rules above and below, set apart from the body.
    /// </summary>
    IntenseQuote,

    /// <summary>
    /// The label under a figure or table.
    /// </summary>
    Caption,

    /// <summary>
    /// Body text inside a page header. Carries the tab stops that centre and right-align.
    /// </summary>
    Header,

    /// <summary>
    /// Body text inside a page footer, with the same tab stops as <see cref="Header"/>.
    /// </summary>
    Footer,

    /// <summary>
    /// The style Word puts on a paragraph it has made part of a list.
    /// </summary>
    ListParagraph,

    /// <summary>
    /// Body text with the space before and after removed.
    /// </summary>
    NoSpacing,

    /// <summary>
    /// The character style that makes a link look like one. Without it a link renders as
    /// ordinary text and only reveals itself on hover.
    /// </summary>
    Hyperlink,

    /// <summary>
    /// Body text inside a footnote — smaller than the body, with no space around it.
    /// </summary>
    FootnoteText,

    /// <summary>
    /// The character style for the superscript mark, both in the body and on the note.
    /// </summary>
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
