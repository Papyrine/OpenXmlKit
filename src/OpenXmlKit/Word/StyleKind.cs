namespace OpenXmlKit.Word;

/// <summary>
/// What a style can be applied to.
/// </summary>
public enum StyleKind
{
    /// <summary>
    /// Applies to a whole paragraph, and carries both paragraph and character formatting.
    /// </summary>
    Paragraph,

    /// <summary>
    /// Applies to a run inside a paragraph, and carries character formatting only.
    /// </summary>
    Character,

    /// <summary>
    /// Applies to a table, and is mostly its conditional blocks. See <see cref="TableStyleConditional"/>.
    /// </summary>
    Table,

    /// <summary>
    /// Ties a paragraph style to a list definition.
    /// </summary>
    Numbering
}
