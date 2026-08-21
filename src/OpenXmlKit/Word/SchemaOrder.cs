using System.Collections.Concurrent;

namespace OpenXmlKit.Word;

/// <summary>
/// Puts a child element where its parent's schema says it goes.
/// </summary>
/// <remarks>
/// Word treats a properties element whose children are out of their <c>CT_*</c> sequence as a
/// corrupt document: it offers to repair on open, and repairing strips the formatting. Almost
/// everywhere, the answer is to assign through the SDK's typed properties, which place a child at
/// its position — see SchemaOrderTests. This exists for the places that cannot.
/// <para>
/// There are two of them. A container the SDK models as a choice rather than a sequence gets no
/// typed children at all: <c>CT_SectPr</c> is one, because header and footer references repeat.
/// And the generator stops emitting typed properties partway through a long sequence, so
/// <c>CT_Settings</c> has one hundred and three children and typed properties for the first
/// twenty-one — which is why a caller wanting <c>w:documentProtection</c> has nothing to assign
/// through and ends up hand-maintaining a list of everything that precedes it.
/// </para>
/// <para>
/// The sequences come from the SDK's own schema data rather than from anyone's reading of the
/// specification, so the two cannot disagree.
/// </para>
/// </remarks>
public static partial class SchemaOrder
{
    static readonly ConcurrentDictionary<string, string[]?> cache = new(StringComparer.Ordinal);

    /// <summary>
    /// Inserts <paramref name="child"/> into <paramref name="parent"/> at its position in the
    /// schema sequence, replacing any element already there under the same name.
    /// </summary>
    /// <remarks>
    /// A child the sequence does not mention — an extension element, or one from a container this
    /// has no entry for — is appended, which is the same thing appending has always done and no
    /// worse than it.
    /// </remarks>
    public static void Place(OpenXmlCompositeElement parent, OpenXmlElement child)
    {
        foreach (var existing in parent.ChildElements.ToList())
        {
            if (existing.LocalName == child.LocalName &&
                existing.NamespaceUri == child.NamespaceUri)
            {
                existing.Remove();
            }
        }

        var sequence = SequenceFor(parent);
        var position = sequence == null ? -1 : Array.IndexOf(sequence, Key(child));
        if (position < 0)
        {
            parent.AppendChild(child);
            return;
        }

        foreach (var sibling in parent.ChildElements)
        {
            // Siblings the sequence does not mention are stepped over rather than displaced: an
            // element with no place in the sequence has no position to be measured against.
            var siblingPosition = Array.IndexOf(sequence!, Key(sibling));
            if (siblingPosition >= 0 &&
                siblingPosition > position)
            {
                parent.InsertBefore(child, sibling);
                return;
            }
        }

        parent.AppendChild(child);
    }

    /// <summary>
    /// Where a child sits in its parent's schema sequence, or -1 when the sequence does not
    /// mention it — which includes every child of a container modelled as a choice, since a choice
    /// imposes no order to have a position in.
    /// </summary>
    public static int IndexOf(OpenXmlElement parent, OpenXmlElement child)
    {
        var sequence = SequenceFor(parent);
        return sequence == null ? -1 : Array.IndexOf(sequence, Key(child));
    }

    static string[]? SequenceFor(OpenXmlElement parent) =>
        cache.GetOrAdd(
            parent.GetType().Name,
            static name => sequences.TryGetValue(name, out var packed) ? packed.Split(' ') : null);

    // The table is keyed the way the schema writes an element name, so a child has to be turned
    // back into that form. An unrecognised namespace keeps the bare local name, which will simply
    // not match — the right outcome for something the sequence does not cover.
    static string Key(OpenXmlElement element) =>
        prefixes.TryGetValue(element.NamespaceUri, out var prefix)
            ? prefix + ":" + element.LocalName
            : element.LocalName;
}
