namespace OpenXmlKit.Word;

/// <summary>
/// The part of a table a conditional style override applies to.
/// </summary>
/// <remarks>
/// Which of these actually take effect is decided by the table using the style, through
/// <see cref="TableLook"/> — a style can define banding that a table then declines to show.
/// </remarks>
public enum TableStyleArea
{
    /// <summary>
    /// Every cell, before any of the more specific areas are laid over it.
    /// </summary>
    WholeTable,

    /// <summary>
    /// The header row.
    /// </summary>
    FirstRow,

    /// <summary>
    /// The total row.
    /// </summary>
    LastRow,

    /// <summary>
    /// The leftmost column, which often carries the row labels.
    /// </summary>
    FirstColumn,

    /// <summary>
    /// The rightmost column.
    /// </summary>
    LastColumn,

    /// <summary>
    /// The odd-numbered horizontal bands — the first stripe of row banding.
    /// </summary>
    Band1Horizontal,

    /// <summary>
    /// The even-numbered horizontal bands.
    /// </summary>
    Band2Horizontal,

    /// <summary>
    /// The odd-numbered vertical bands, the first stripe of column banding.
    /// </summary>
    Band1Vertical,

    /// <summary>
    /// The even-numbered vertical bands.
    /// </summary>
    Band2Vertical,

    /// <summary>
    /// The single cell where the first row meets the first column.
    /// </summary>
    TopLeftCell,

    /// <summary>
    /// The cell where the first row meets the last column.
    /// </summary>
    TopRightCell,

    /// <summary>
    /// The cell where the last row meets the first column.
    /// </summary>
    BottomLeftCell,

    /// <summary>
    /// The cell where the last row meets the last column.
    /// </summary>
    BottomRightCell
}
