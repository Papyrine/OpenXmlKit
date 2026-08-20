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

    FirstRow,
    LastRow,
    FirstColumn,
    LastColumn,

    /// <summary>
    /// The odd-numbered horizontal bands — the first stripe of row banding.
    /// </summary>
    Band1Horizontal,

    /// <summary>
    /// The even-numbered horizontal bands.
    /// </summary>
    Band2Horizontal,

    Band1Vertical,
    Band2Vertical,

    /// <summary>
    /// The single cell where the first row meets the first column.
    /// </summary>
    TopLeftCell,

    TopRightCell,
    BottomLeftCell,
    BottomRightCell
}
