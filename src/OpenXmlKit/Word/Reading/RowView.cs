namespace OpenXmlKit.Word;

/// <summary>
/// A table row in a document being read.
/// </summary>
public readonly struct RowView
{
    readonly W.TableRow element;

    internal RowView(W.TableRow element) =>
        this.element = element;

    /// <summary>
    /// The row's text, cells joined by tabs.
    /// </summary>
    public string Text =>
        string.Join("\t", Cells.Select(_ => _.Text));

    /// <summary>
    /// The cells in the row, left to right.
    /// </summary>
    public IEnumerable<CellView> Cells
    {
        get
        {
            foreach (var cell in element.Elements<W.TableCell>())
            {
                yield return new(cell);
            }
        }
    }

    /// <summary>
    /// The formatting written on the row.
    /// </summary>
    public IRowFormatView Format
    {
        get
        {
            var format = new RowFormat();
            format.ReadFrom(element.TableRowProperties);
            return format;
        }
    }

    /// <summary>
    /// Whether this row repeats at the top of every page the table spans.
    /// </summary>
    public bool IsHeader =>
        element.TableRowProperties?.GetFirstChild<W.TableHeader>() != null;

    /// <summary>
    /// The underlying OpenXML element, for anything this view does not expose.
    /// </summary>
    public W.TableRow ToOpenXml() =>
        element;
}
