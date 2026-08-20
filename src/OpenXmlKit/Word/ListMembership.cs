namespace OpenXmlKit.Word;

/// <summary>
/// A paragraph's place in a list: which numbering definition it belongs to, and how deep.
/// </summary>
/// <param name="NumberingId">
/// The numbering instance, as allocated by <see cref="Numbering"/>.
/// </param>
/// <param name="Level">
/// Depth, from 0 for the outermost. Word supports nine.
/// </param>
public readonly record struct ListMembership(int NumberingId, int Level);
