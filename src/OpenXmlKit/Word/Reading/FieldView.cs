namespace OpenXmlKit.Word;

/// <summary>
/// A field in a document being read — a value Word computes, such as a page number or a
/// cross-reference.
/// </summary>
/// <remarks>
/// A record rather than a view over an element, because the usual form of a field is not one
/// element: it is a begin marker, an instruction, a separator, the cached result and an end
/// marker, spread across five runs. What a caller wants from it is the two strings.
/// </remarks>
/// <param name="Code">
/// The field instruction — <c>PAGE</c>, <c>PAGEREF du3 \h</c>.
/// </param>
/// <param name="Value">
/// The cached result: what Word shows until it recalculates.
/// </param>
public readonly record struct FieldView(string Code, string Value)
{
    /// <summary>
    /// A readable form, for logs and debugging rather than for the file.
    /// </summary>
    public override string ToString() =>
        Code;
}
