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
    /// The cell's text, blocks joined by newlines, in document order. A nested table is included,
    /// on the same terms as <see cref="BlockContainerView.Text"/>.
    /// </summary>
    public string Text =>
        new BlockContainerView(element).Text;

    /// <summary>
    /// The paragraphs in the cell.
    /// </summary>
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

    /// <summary>
    /// The formatting written on the cell.
    /// </summary>
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

    /// <summary>
    /// A readable form, for logs and debugging rather than for the file.
    /// </summary>
    public override string ToString() =>
        Text;
}
