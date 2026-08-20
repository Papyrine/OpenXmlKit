namespace OpenXmlKit.Word;

/// <summary>
/// Which conditional parts of a table style apply to this table.
/// </summary>
/// <remarks>
/// A table style can define a distinct look for the header row, the total row, the first and last
/// columns, and alternating bands. This is the switchboard that says which of those Word should
/// draw — a style with header formatting produces a plain table if <see cref="FirstRow"/> is off.
/// <para>
/// The defaults match what Word writes when a table is inserted from the ribbon: header row on,
/// row banding on, everything else off.
/// </para>
/// </remarks>
public class TableLook :
    ITableLookView
{
    /// <summary>
    /// Whether the header-row block of the table style applies.
    /// </summary>
    public bool FirstRow { get; set; } = true;

    /// <summary>
    /// Whether the total-row block applies.
    /// </summary>
    public bool LastRow { get; set; }

    /// <summary>
    /// Whether the first-column block applies.
    /// </summary>
    public bool FirstColumn { get; set; }

    /// <summary>
    /// Whether the last-column block applies.
    /// </summary>
    public bool LastColumn { get; set; }

    /// <summary>
    /// Alternating shading down the rows.
    /// </summary>
    public bool RowBanding { get; set; } = true;

    /// <summary>
    /// Alternating shading across the columns.
    /// </summary>
    public bool ColumnBanding { get; set; }

    /// <summary>
    /// Whether this is what Word writes for a freshly inserted table, and so says nothing a
    /// reader could not assume.
    /// </summary>
    internal bool IsDefault =>
        FirstRow &&
        !LastRow &&
        !FirstColumn &&
        !LastColumn &&
        RowBanding &&
        !ColumnBanding;

    /// <summary>
    /// Copies every flag from the other look.
    /// </summary>
    public void CopyFrom(TableLook other)
    {
        FirstRow = other.FirstRow;
        LastRow = other.LastRow;
        FirstColumn = other.FirstColumn;
        LastColumn = other.LastColumn;
        RowBanding = other.RowBanding;
        ColumnBanding = other.ColumnBanding;
    }

    internal W.TableLook ToOpenXml() =>
        new()
        {
            // The hex val is the legacy form of the same six flags, and Word still reads it in
            // preference to the attributes on some paths, so both are written.
            Val = Hex(),
            FirstRow = FirstRow,
            LastRow = LastRow,
            FirstColumn = FirstColumn,
            LastColumn = LastColumn,
            // The banding attributes are stated as suppression rather than as application.
            NoHorizontalBand = !RowBanding,
            NoVerticalBand = !ColumnBanding
        };

    string Hex()
    {
        var value = 0;
        if (FirstRow)
        {
            value |= 0x0020;
        }

        if (LastRow)
        {
            value |= 0x0040;
        }

        if (FirstColumn)
        {
            value |= 0x0080;
        }

        if (LastColumn)
        {
            value |= 0x0100;
        }

        if (!RowBanding)
        {
            value |= 0x0200;
        }

        if (!ColumnBanding)
        {
            value |= 0x0400;
        }

        return value.ToString("X4", CultureInfo.InvariantCulture);
    }

    internal void ReadFrom(W.TableLook? look)
    {
        if (look == null)
        {
            return;
        }

        FirstRow = look.FirstRow?.Value ?? FirstRow;
        LastRow = look.LastRow?.Value ?? LastRow;
        FirstColumn = look.FirstColumn?.Value ?? FirstColumn;
        LastColumn = look.LastColumn?.Value ?? LastColumn;

        if (look.NoHorizontalBand?.Value is { } noHorizontal)
        {
            RowBanding = !noHorizontal;
        }

        if (look.NoVerticalBand?.Value is { } noVertical)
        {
            ColumnBanding = !noVertical;
        }
    }
}
