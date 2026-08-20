namespace OpenXmlKit.Word;

/// <summary>
/// A cell's part in a vertical merge. Horizontal merging is a span count rather than a state —
/// see <c>CellFormat.ColumnSpan</c>.
/// </summary>
public enum CellMerge
{
    /// <summary>
    /// Not merged.
    /// </summary>
    None,

    /// <summary>
    /// The top cell of a merged run, and the one whose content shows.
    /// </summary>
    Restart,

    /// <summary>
    /// Absorbed into the cell above.
    /// </summary>
    Continue
}
