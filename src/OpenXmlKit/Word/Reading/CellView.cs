namespace OpenXmlKit.Word;

/// <summary>
/// A table cell in a document being read.
/// </summary>
public readonly struct CellView
{
    readonly W.TableCell element;

    internal CellView(W.TableCell element) =>
        this.element = element;

    /// <summary>
    /// The cell's text, paragraphs joined by newlines.
    /// </summary>
    public string Text =>
        string.Join("\n", Paragraphs.Select(_ => _.Text));

    public IEnumerable<ParagraphView> Paragraphs
    {
        get
        {
            foreach (var paragraph in element.Elements<W.Paragraph>())
            {
                yield return new(paragraph);
            }
        }
    }

    /// <summary>
    /// Tables nested inside this cell.
    /// </summary>
    public IEnumerable<TableView> Tables
    {
        get
        {
            foreach (var table in element.Elements<W.Table>())
            {
                yield return new(table);
            }
        }
    }

    public ICellFormatView Format
    {
        get
        {
            var format = new CellFormat();
            format.ReadFrom(element.TableCellProperties);
            return format;
        }
    }

    /// <summary>
    /// How many grid columns this cell covers.
    /// </summary>
    public int ColumnSpan =>
        element.TableCellProperties?.GridSpan?.Val?.Value ?? 1;

    /// <summary>
    /// The underlying OpenXML element, for anything this view does not expose.
    /// </summary>
    public W.TableCell ToOpenXml() =>
        element;

    public override string ToString() =>
        Text;
}
