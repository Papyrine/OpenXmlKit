namespace OpenXmlKit.Word;

/// <summary>
/// Which conditional parts of a table style apply, as read.
/// </summary>
public interface ITableLookView
{
    bool FirstRow { get; }
    bool LastRow { get; }
    bool FirstColumn { get; }
    bool LastColumn { get; }
    bool RowBanding { get; }
    bool ColumnBanding { get; }
}
