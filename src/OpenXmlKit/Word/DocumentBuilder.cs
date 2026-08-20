namespace OpenXmlKit.Word;

/// <summary>
/// A cursor over a document, for writing it a statement at a time.
/// </summary>
/// <remarks>
/// The other way to build content is the nested one — <see cref="Table.Create"/> and the
/// lambda-taking methods on <see cref="Body"/>, <see cref="Paragraph"/> and the rest — which suits
/// a self-contained fragment. This suits a document written front to back, where the shape of the
/// code should follow the shape of the prose rather than nest nine deep.
/// <para>
/// Formatting is ambient: set <see cref="Font"/> or <see cref="ParagraphFormat"/> and everything
/// written afterwards carries it, until it is changed back or a
/// <see cref="PushFormatting"/> scope ends. Unlike the model this borrows from, the scope covers
/// character, paragraph, cell, row and table formatting rather than character formatting alone — a
/// scope that restores only some of what it saved is a trap rather than a convenience.
/// </para>
/// <para>
/// Ambient formatting is captured when an element closes rather than when it opens, so a style
/// named part-way through a paragraph, or a width set inside a table scope, still applies to the
/// element it was named in.
/// </para>
/// </remarks>
public class DocumentBuilder
{
    readonly Document document;
    readonly Stack<BlockContainer> containers = new();
    readonly Stack<Formatting> formatting = new();

    Paragraph? paragraph;
    Table? table;
    Row? row;
    Cell? cell;

    internal DocumentBuilder(Document document)
    {
        this.document = document;
        containers.Push(document.Body);
    }

    /// <summary>
    /// Character formatting applied to everything written from here.
    /// </summary>
    public Font Font { get; private set; } = new();

    /// <summary>
    /// Paragraph formatting applied to every paragraph written from here.
    /// </summary>
    public ParagraphFormat ParagraphFormat { get; private set; } = new();

    /// <summary>
    /// Cell formatting applied to every cell opened from here.
    /// </summary>
    public CellFormat CellFormat { get; private set; } = new();

    /// <summary>
    /// Row formatting applied to every row opened from here.
    /// </summary>
    public RowFormat RowFormat { get; private set; } = new();

    /// <summary>
    /// Table formatting applied to every table opened from here.
    /// </summary>
    public TableFormat TableFormat { get; private set; } = new();

    /// <summary>
    /// The paragraph being written into, if any.
    /// </summary>
    public Paragraph? CurrentParagraph => paragraph;

    /// <summary>
    /// Whether nothing has been written into the current paragraph yet.
    /// </summary>
    public bool IsAtStartOfParagraph => paragraph == null;

    /// <summary>
    /// Saves every formatting object, restoring them all when the returned scope is disposed.
    /// </summary>
    public IDisposable PushFormatting()
    {
        formatting.Push(new(Font, ParagraphFormat, CellFormat, RowFormat, TableFormat));
        Font = Font.Clone();
        ParagraphFormat = ParagraphFormat.Clone();
        CellFormat = CellFormat.Clone();
        RowFormat = RowFormat.Clone();
        TableFormat = TableFormat.Clone();
        return new Scope(PopFormatting);
    }

    void PopFormatting()
    {
        if (formatting.Count == 0)
        {
            return;
        }

        var saved = formatting.Pop();
        Font = saved.Font;
        ParagraphFormat = saved.ParagraphFormat;
        CellFormat = saved.CellFormat;
        RowFormat = saved.RowFormat;
        TableFormat = saved.TableFormat;
    }

    readonly record struct Formatting(
        Font Font,
        ParagraphFormat ParagraphFormat,
        CellFormat CellFormat,
        RowFormat RowFormat,
        TableFormat TableFormat);

    /// <summary>
    /// Writes text into the current paragraph, starting one if there is none.
    /// </summary>
    public DocumentBuilder Write(string? text)
    {
        var run = EnsureParagraph().AddRun(text);
        run.Font.CopyFrom(Font);
        return this;
    }

    /// <summary>
    /// Writes text and ends the paragraph.
    /// </summary>
    public DocumentBuilder Writeln(string? text = null)
    {
        Write(text);
        EndParagraph();
        return this;
    }

    /// <summary>
    /// Ends the current paragraph, so the next write starts a new one.
    /// </summary>
    public DocumentBuilder EndParagraph()
    {
        if (paragraph == null)
        {
            // An explicit blank line, rather than a no-op: a caller asking to end a paragraph that
            // was never started means to leave an empty one behind.
            EnsureParagraph();
        }

        FlushParagraph();
        return this;
    }

    /// <summary>
    /// Inserts a break. A page or column break ends the paragraph as well.
    /// </summary>
    public DocumentBuilder InsertBreak(BreakKind kind = BreakKind.Line)
    {
        var run = EnsureParagraph().AddRun();
        run.Font.CopyFrom(Font);
        run.AppendBreak(kind);
        if (kind != BreakKind.Line)
        {
            FlushParagraph();
        }

        return this;
    }

    /// <summary>
    /// Inserts an element this library does not model into the current paragraph.
    /// </summary>
    public DocumentBuilder InsertElement(OpenXmlElement element)
    {
        EnsureParagraph().AppendElement(element);
        return this;
    }

    /// <summary>
    /// Applies a paragraph style to what is written from here.
    /// </summary>
    public DocumentBuilder Style(string? styleId)
    {
        ParagraphFormat.StyleId = styleId;
        return this;
    }

    /// <summary>
    /// Applies a built-in paragraph style, writing its definition into the document if it is not
    /// there already.
    /// </summary>
    public DocumentBuilder Style(BuiltInStyle style)
    {
        var definition = document.Styles.EnsureBuiltIn(style);
        ParagraphFormat.StyleId = definition.Id;
        return this;
    }

    /// <summary>
    /// Writes a paragraph in a built-in heading style.
    /// </summary>
    public DocumentBuilder Heading(int level, string text)
    {
        if (level is < 1 or > 6)
        {
            throw new ArgumentOutOfRangeException(nameof(level), level, "Heading levels run from 1 to 6.");
        }

        using (PushFormatting())
        {
            Style((BuiltInStyle) ((int) BuiltInStyle.Heading1 + level - 1));
            Writeln(text);
        }

        return this;
    }

    /// <summary>
    /// Makes the paragraphs written from here items of a list.
    /// </summary>
    public DocumentBuilder ListItem(ListDefinition list, int depth = 0)
    {
        ParagraphFormat.List = list.At(depth);
        return this;
    }

    /// <summary>
    /// Stops the paragraphs written from here being list items.
    /// </summary>
    public DocumentBuilder EndList()
    {
        ParagraphFormat.List = null;
        return this;
    }

    /// <summary>
    /// Starts a table, and returns a scope that ends it.
    /// </summary>
    /// <remarks>
    /// The scope exists so that a table cannot be left open: paired start and end calls can be
    /// unbalanced, a using block cannot.
    /// </remarks>
    public IDisposable Table()
    {
        StartTable();
        return new Scope(() => EndTable());
    }

    /// <summary>
    /// Starts a table.
    /// </summary>
    public Table StartTable()
    {
        FlushParagraph();
        var started = new Table();
        table = started;
        return started;
    }

    /// <summary>
    /// Starts a row, and returns a scope that ends it.
    /// </summary>
    public IDisposable Row()
    {
        StartRow();
        return new Scope(() => EndRow());
    }

    /// <summary>
    /// Starts a row.
    /// </summary>
    public Row StartRow()
    {
        var started = RequireTable().AddRow();
        row = started;
        return started;
    }

    /// <summary>
    /// Starts a cell, ending the previous one.
    /// </summary>
    public Cell InsertCell()
    {
        // The previous cell has to be closed first, or its container stays on the stack and the
        // table ends up being appended into one of its own cells — a cycle, which the SDK walks
        // until the stack runs out rather than rejecting.
        EndCell();
        var started = (row ?? StartRow()).AddCell();
        cell = started;
        containers.Push(new CellContainer(started));
        return started;
    }

    /// <summary>
    /// Ends the current row.
    /// </summary>
    public Row? EndRow()
    {
        EndCell();
        var ended = row;
        row = null;
        ended?.Format.CopyFrom(RowFormat);
        return ended;
    }

    /// <summary>
    /// Ends the current table and appends it to whatever contains it.
    /// </summary>
    public Table? EndTable()
    {
        EndRow();
        var ended = table;
        table = null;
        if (ended != null)
        {
            ended.Format.CopyFrom(TableFormat);
            Containers.Peek().Append(ended);
        }

        return ended;
    }

    void EndCell()
    {
        if (cell == null)
        {
            return;
        }

        FlushParagraph();
        cell.Format.CopyFrom(CellFormat);
        cell = null;
        if (containers.Peek() is CellContainer)
        {
            containers.Pop();
        }
    }

    Stack<BlockContainer> Containers => containers;

    Table RequireTable() =>
        table ?? StartTable();

    /// <summary>
    /// Moves the cursor into a header, so what is written next goes there.
    /// </summary>
    public DocumentBuilder MoveToHeader(HeaderFooterKind kind = HeaderFooterKind.Default)
    {
        FlushParagraph();
        containers.Push(document.Body.Section.AddHeader(kind));
        return this;
    }

    /// <summary>
    /// Moves the cursor into a footer.
    /// </summary>
    public DocumentBuilder MoveToFooter(HeaderFooterKind kind = HeaderFooterKind.Default)
    {
        FlushParagraph();
        containers.Push(document.Body.Section.AddFooter(kind));
        return this;
    }

    /// <summary>
    /// Returns the cursor to the document body.
    /// </summary>
    public DocumentBuilder MoveToBody()
    {
        FlushParagraph();
        while (containers.Count > 1)
        {
            containers.Pop();
        }

        return this;
    }

    Paragraph EnsureParagraph()
    {
        if (paragraph != null)
        {
            return paragraph;
        }

        paragraph = new();
        Containers.Peek().Append(paragraph);
        return paragraph;
    }

    // The ambient paragraph formatting is copied in when the paragraph closes rather than when it
    // opens, so a style named part-way through still applies to the paragraph it was named in.
    void FlushParagraph()
    {
        if (paragraph == null)
        {
            return;
        }

        paragraph.Format.CopyFrom(ParagraphFormat);
        paragraph.MarkFont.CopyFrom(Font);
        _ = paragraph.ToOpenXml();
        paragraph = null;
    }

    internal void Flush()
    {
        FlushParagraph();
        if (table != null)
        {
            EndTable();
        }
    }

    sealed class Scope(Action onDispose) :
        IDisposable
    {
        bool disposed;

        public void Dispose()
        {
            if (disposed)
            {
                return;
            }

            disposed = true;
            onDispose();
        }
    }

    // A cell is a block container like the body is, but it is not one in the object model, so this
    // adapts it for the container stack.
    sealed class CellContainer(Cell cell) :
        BlockContainer(cell.Container, null)
    {
        public Cell Cell { get; } = cell;
    }
}
