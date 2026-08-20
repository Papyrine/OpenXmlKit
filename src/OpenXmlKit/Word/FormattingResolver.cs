namespace OpenXmlKit.Word;

/// <summary>
/// Works out the formatting that actually applies to a piece of content, rather than the formatting
/// written on it.
/// </summary>
/// <remarks>
/// Word resolves formatting through a cascade, and most of the surprises in building documents come
/// from not knowing its order. The one that catches everybody: a paragraph style outranks a table
/// style, so branding a table's font through its table style alone silently loses to whatever
/// Normal says, and the fix is a paragraph style on the cells rather than more table styling.
/// <para>
/// The order applied here, lowest first, follows the precedence the format defines:
/// document defaults, table style, numbering style, paragraph style, character style, then the
/// direct formatting on the paragraph and finally on the run. Each style contributes through its
/// own <c>basedOn</c> chain, walked from the root down so that a derived style overlays the one it
/// derives from.
/// </para>
/// </remarks>
public class FormattingResolver
{
    readonly Document document;

    internal FormattingResolver(Document document) =>
        this.document = document;

    /// <summary>
    /// The character formatting a run ends up with.
    /// </summary>
    /// <param name="run">
    /// The run whose direct formatting sits at the top of the cascade.
    /// </param>
    /// <param name="paragraph">
    /// The paragraph containing it, which contributes its style and its direct formatting. Omitting
    /// it resolves the run against the document defaults alone.
    /// </param>
    /// <param name="tableStyleId">
    /// The style of the table the content sits in, if any.
    /// </param>
    public Font FontFor(Run run, Paragraph? paragraph = null, string? tableStyleId = null)
    {
        var resolved = new Font();

        if (DefaultFont() is { } defaults)
        {
            resolved.MergeFrom(defaults);
        }

        if (tableStyleId != null)
        {
            ApplyStyleFonts(resolved, tableStyleId);
        }

        if (paragraph?.Format.StyleId is { } paragraphStyle)
        {
            ApplyStyleFonts(resolved, paragraphStyle);
        }

        if (run.Font.StyleId is { } characterStyle)
        {
            ApplyStyleFonts(resolved, characterStyle);
        }

        resolved.MergeFrom(run.Font);
        return resolved;
    }

    /// <summary>
    /// The paragraph formatting a paragraph ends up with.
    /// </summary>
    public ParagraphFormat FormatFor(Paragraph paragraph, string? tableStyleId = null)
    {
        var resolved = new ParagraphFormat();

        if (DefaultParagraphFormat() is { } defaults)
        {
            resolved.MergeFrom(defaults);
        }

        if (tableStyleId != null)
        {
            ApplyStyleFormats(resolved, tableStyleId);
        }

        if (paragraph.Format.StyleId is { } styleId)
        {
            ApplyStyleFormats(resolved, styleId);
        }

        resolved.MergeFrom(paragraph.Format);
        // The style id is what the paragraph names, not something the cascade produced, so it is
        // carried through rather than resolved away - a caller asking what applies still wants to
        // know which style said so.
        resolved.StyleId = paragraph.Format.StyleId;
        return resolved;
    }

    // A style contributes what it inherits before what it declares, so the basedOn chain is walked
    // to its root and then applied downwards. A cycle in the chain - which a hand-edited template
    // can carry - would otherwise not terminate.
    IReadOnlyList<Style> Chain(string styleId)
    {
        var chain = new List<Style>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var current = styleId;
        while (current != null && seen.Add(current))
        {
            if (document.Styles.Find(current) is not { } style)
            {
                break;
            }

            chain.Add(style);
            current = style.BasedOn;
        }

        chain.Reverse();
        return chain;
    }

    void ApplyStyleFonts(Font resolved, string styleId)
    {
        foreach (var style in Chain(styleId))
        {
            resolved.MergeFrom(style.Font);
        }
    }

    void ApplyStyleFormats(ParagraphFormat resolved, string styleId)
    {
        foreach (var style in Chain(styleId))
        {
            resolved.MergeFrom(style.ParagraphFormat);
        }
    }

    Font? DefaultFont()
    {
        var defaults = document.MainPart.StyleDefinitionsPart?.Styles?.DocDefaults;
        if (defaults?.RunPropertiesDefault?.RunPropertiesBaseStyle is not { } source)
        {
            return null;
        }

        var properties = new W.RunProperties();
        foreach (var child in source.ChildElements)
        {
            properties.AppendChild(child.CloneNode(true));
        }

        var font = new Font();
        font.ReadFrom(properties);
        return font;
    }

    ParagraphFormat? DefaultParagraphFormat()
    {
        var defaults = document.MainPart.StyleDefinitionsPart?.Styles?.DocDefaults;
        if (defaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle is not { } source)
        {
            return null;
        }

        var properties = new W.ParagraphProperties();
        foreach (var child in source.ChildElements)
        {
            properties.AppendChild(child.CloneNode(true));
        }

        var format = new ParagraphFormat();
        format.ReadFrom(properties);
        return format;
    }
}
