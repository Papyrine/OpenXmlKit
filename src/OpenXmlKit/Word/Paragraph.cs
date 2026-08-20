namespace OpenXmlKit.Word;

/// <summary>
/// A paragraph, and the runs it contains.
/// </summary>
/// <remarks>
/// Two dialects are offered for building one, and they compose. The <c>Add</c> methods return what
/// they created, for when the caller wants to keep working on it; the lambda-taking methods return
/// the paragraph, for when the caller wants to keep chaining. Both reach the same place.
/// </remarks>
public partial class Paragraph
{
    readonly W.Paragraph element;
    ParagraphFormat? format;
    Font? markFont;

    // The runs this paragraph created. Their elements are already in the tree; what is kept here is
    // the wrapper, so that formatting set on a run after it was added still reaches the element
    // when the tree is flushed. Everything in the object model owns its children this way.
    readonly List<Run> runs = [];

    /// <summary>
    /// A paragraph with no content and no formatting.
    /// </summary>
    public Paragraph() =>
        element = new();

    /// <summary>
    /// A paragraph holding one run of plain text.
    /// </summary>
    public Paragraph(string? text)
        : this()
    {
        if (!string.IsNullOrEmpty(text))
        {
            AddRun(text);
        }
    }

    /// <summary>
    /// Paragraph formatting.
    /// </summary>
    public ParagraphFormat Format => format ??= new();

    /// <summary>
    /// Formatting for the paragraph mark itself, which decides the height of an empty paragraph.
    /// </summary>
    public Font MarkFont => markFont ??= new();

    /// <summary>
    /// Adds a run, optionally with text, and returns it.
    /// </summary>
    public Run AddRun(string? text = null)
    {
        var run = new Run();
        run.Append(text);
        element.AppendChild(run.Element);
        runs.Add(run);
        return run;
    }

    /// <summary>
    /// Adds a run and configures it, returning the paragraph so the call can be chained.
    /// </summary>
    public Paragraph AddRun(Action<Run> configure)
    {
        configure(AddRun());
        return this;
    }

    /// <summary>
    /// Adds a run of text, optionally with its own character formatting.
    /// </summary>
    public Paragraph Append(string? text, Action<Font>? font = null)
    {
        var run = AddRun(text);
        font?.Invoke(run.Font);
        return this;
    }

    /// <summary>
    /// Appends a run of bold text.
    /// </summary>
    public Paragraph Bold(string? text)
    {
        AddRun(text).Bold();
        return this;
    }

    /// <summary>
    /// Appends a run of italic text.
    /// </summary>
    public Paragraph Italic(string? text)
    {
        AddRun(text).Italic();
        return this;
    }

    /// <summary>
    /// Applies a paragraph style by id.
    /// </summary>
    public Paragraph Style(string? styleId)
    {
        Format.StyleId = styleId;
        return this;
    }

    /// <summary>
    /// Sets how the paragraph sits between its margins.
    /// </summary>
    public Paragraph Alignment(ParagraphAlignment alignment)
    {
        Format.Alignment = alignment;
        return this;
    }

    /// <summary>
    /// Configures the paragraph formatting.
    /// </summary>
    public Paragraph Formatting(Action<ParagraphFormat> configure)
    {
        configure(Format);
        return this;
    }

    /// <summary>
    /// Adds a break in a run of its own.
    /// </summary>
    public Paragraph AppendBreak(BreakKind kind = BreakKind.Line)
    {
        AddRun().AppendBreak(kind);
        return this;
    }

    /// <summary>
    /// Appends an element this library does not model — a field, a drawing, a content control.
    /// </summary>
    public Paragraph AppendElement(OpenXmlElement child)
    {
        element.AppendChild(child);
        return this;
    }

    /// <summary>
    /// The underlying OpenXML element, with any pending content and formatting applied.
    /// </summary>
    public W.Paragraph ToOpenXml()
    {
        Flush();
        return element;
    }

    internal W.Paragraph Element => element;

    // Content is appended as it is added, so this only rebuilds the properties. pPr has to be the
    // paragraph's first child, and assigning the typed property is what puts it there however many
    // runs are already in place - which is also why content needs no buffering to stay behind it.
    internal void Flush()
    {
        foreach (var run in runs)
        {
            run.Flush();
        }

        if (format == null &&
            markFont == null)
        {
            return;
        }

        element.ParagraphProperties = (format ?? new()).ToProperties(markFont);
    }
}
