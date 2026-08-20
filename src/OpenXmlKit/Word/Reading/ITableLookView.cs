namespace OpenXmlKit.Word;

/// <summary>
/// Which conditional parts of a table style apply, as read.
/// </summary>
public interface ITableLookView
{
    /// <inheritdoc cref="TableLook.FirstRow"/>
    bool FirstRow { get; }

    /// <inheritdoc cref="TableLook.LastRow"/>
    bool LastRow { get; }

    /// <inheritdoc cref="TableLook.FirstColumn"/>
    bool FirstColumn { get; }

    /// <inheritdoc cref="TableLook.LastColumn"/>
    bool LastColumn { get; }

    /// <inheritdoc cref="TableLook.RowBanding"/>
    bool RowBanding { get; }

    /// <inheritdoc cref="TableLook.ColumnBanding"/>
    bool ColumnBanding { get; }
}
