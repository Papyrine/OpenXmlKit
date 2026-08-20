namespace OpenXmlKit.Word;

/// <summary>
/// The list definitions a document being read carries.
/// </summary>
/// <remarks>
/// A paragraph states only which numbering instance it belongs to and how deep — reading
/// <see cref="ParagraphView.List"/> tells you a paragraph is in list 3 at level 1, and nothing
/// about whether that renders as a bullet or as "b.". Resolving the one to the other is what this
/// is for.
/// <para>
/// A numbering instance may override individual levels of the definition it points at. Those
/// overrides are not applied here, so a document that uses them reports the underlying definition
/// rather than the effective one.
/// </para>
/// </remarks>
public sealed class NumberingView :
    IEnumerable<ListDefinitionView>
{
    readonly W.Numbering? root;

    internal NumberingView(MainDocumentPart part) =>
        root = part.NumberingDefinitionsPart?.Numbering;

    /// <summary>
    /// The definition a numbering id resolves to, or null when the document has none.
    /// </summary>
    public ListDefinitionView? this[int numberingId] => Find(numberingId);

    /// <summary>
    /// The definition a numbering id resolves to, or null when the document has none.
    /// </summary>
    public ListDefinitionView? Find(int numberingId)
    {
        if (root == null)
        {
            return null;
        }

        foreach (var instance in root.Elements<W.NumberingInstance>())
        {
            if (instance.NumberID?.Value != numberingId ||
                instance.AbstractNumId?.Val?.Value is not { } abstractId)
            {
                continue;
            }

            foreach (var definition in root.Elements<W.AbstractNum>())
            {
                if (definition.AbstractNumberId?.Value == abstractId)
                {
                    return new(numberingId, definition);
                }
            }
        }

        return null;
    }

    /// <summary>
    /// The level a paragraph's list membership resolves to — the marker it actually draws.
    /// </summary>
    public IListLevelView? LevelFor(ListMembership membership) =>
        Find(membership.NumberingId)?[membership.Level];

    /// <summary>
    /// Every list definition the document carries.
    /// </summary>
    public IEnumerator<ListDefinitionView> GetEnumerator()
    {
        if (root == null)
        {
            yield break;
        }

        foreach (var instance in root.Elements<W.NumberingInstance>())
        {
            if (instance.NumberID?.Value is { } id &&
                Find(id) is { } definition)
            {
                yield return definition;
            }
        }
    }

    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();
}
