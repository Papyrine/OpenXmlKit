using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using W = DocumentFormat.OpenXml.Wordprocessing;

BenchmarkRunner.Run<TableBenchmarks>();

/// <summary>
/// The wrapper against the raw SDK, on the shape the estate actually builds: a table of rows, each
/// cell a paragraph of text.
/// </summary>
/// <remarks>
/// The point is not to win. The wrapper does strictly more than the raw path below — it fills the
/// grid, writes the required tblPr, and orders every properties element — so the number to watch is
/// how much that costs, not whether it costs anything.
/// </remarks>
[MemoryDiagnoser]
public class TableBenchmarks
{
    [Params(10, 200)]
    public int Rows { get; set; }

    [Benchmark(Baseline = true)]
    public byte[] RawSdk()
    {
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            var body = new W.Body();

            var table = new W.Table();
            table.Append(
                new W.TableProperties(
                    new W.TableWidth
                    {
                        Width = "5000",
                        Type = W.TableWidthUnitValues.Pct
                    }));
            var grid = new W.TableGrid();
            grid.Append(new W.GridColumn());
            grid.Append(new W.GridColumn());
            table.Append(grid);

            for (var index = 0; index < Rows; index++)
            {
                var row = new W.TableRow();
                for (var column = 0; column < 2; column++)
                {
                    row.Append(
                        new W.TableCell(
                            new W.Paragraph(
                                new W.Run(
                                    new W.RunProperties(new W.Bold()),
                                    new W.Text($"cell {index}.{column}")
                                    {
                                        Space = SpaceProcessingModeValues.Preserve
                                    }))));
                }

                table.Append(row);
            }

            body.Append(table);
            body.Append(new W.Paragraph());
            main.Document = new(body);
        }

        return stream.ToArray();
    }

    [Benchmark]
    public byte[] OpenXmlKit()
    {
        using var document = Document.Create();
        var table = document.Body.AddTable();
        table.Format.Width = Width.Percent(100);

        for (var index = 0; index < Rows; index++)
        {
            var row = table.AddRow();
            for (var column = 0; column < 2; column++)
            {
                var cell = row.AddCell();
                cell.AddParagraph().AddRun($"cell {index}.{column}").Bold();
            }
        }

        return document.ToArray();
    }
}
