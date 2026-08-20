namespace OpenXmlKit.Word;

/// <summary>
/// A run of pages sharing one page setup and one set of headers and footers.
/// </summary>
public class Section
{
    readonly W.SectionProperties element;
    readonly Document? document;
    PageSetup? pageSetup;

    internal Section(W.SectionProperties element, Document? document)
    {
        this.element = element;
        this.document = document;
        pageSetup = new();
        pageSetup.ReadFrom(element);
    }

    /// <summary>
    /// Page geometry for this section.
    /// </summary>
    public PageSetup PageSetup => pageSetup ??= new();

    /// <summary>
    /// Where the content after this section break begins.
    /// </summary>
    public SectionStart? Start
    {
        get => PageSetup.Start;
        set => PageSetup.Start = value;
    }

    /// <summary>
    /// Adds or replaces a header for this section, and returns its body to write into.
    /// </summary>
    /// <remarks>
    /// A first-page or even-page header also needs the section to opt into it — see
    /// <see cref="Word.PageSetup.DifferentFirstPage"/> and
    /// <see cref="Word.PageSetup.DifferentOddAndEvenPages"/> — or Word stores it and never shows it.
    /// This sets the matching flag so the header appears, rather than leaving the caller to
    /// discover the omission by its absence.
    /// </remarks>
    public HeaderFooter AddHeader(HeaderFooterKind kind = HeaderFooterKind.Default)
    {
        var main = RequireDocument().MainPart;
        var part = main.AddNewPart<HeaderPart>();
        part.Header = new();
        SetFlags(kind);
        Reference<W.HeaderReference>(kind, main.GetIdOfPart(part));
        return new(part.Header, RequireDocument());
    }

    /// <summary>
    /// Adds or replaces a footer for this section, and returns its body to write into.
    /// </summary>
    public HeaderFooter AddFooter(HeaderFooterKind kind = HeaderFooterKind.Default)
    {
        var main = RequireDocument().MainPart;
        var part = main.AddNewPart<FooterPart>();
        part.Footer = new();
        SetFlags(kind);
        Reference<W.FooterReference>(kind, main.GetIdOfPart(part));
        return new(part.Footer, RequireDocument());
    }

    void SetFlags(HeaderFooterKind kind)
    {
        if (kind == HeaderFooterKind.First)
        {
            PageSetup.DifferentFirstPage = true;
        }
        else if (kind == HeaderFooterKind.Even)
        {
            PageSetup.DifferentOddAndEvenPages = true;
            RequireDocument().EnableEvenAndOddHeaders();
        }
    }

    // Header and footer references have to lead sectPr — Word repairs the document when page size
    // comes first, and a stricter reader rejects it outright. They are repeatable, so the SDK has
    // no typed property to assign through and no ordering to inherit from it; prepending is what
    // puts them in front of whatever page setup has already written.
    void Reference<T>(HeaderFooterKind kind, string relationshipId)
        where T : OpenXmlLeafElement, new()
    {
        var type = kind.ToOpenXml();
        foreach (var existing in element.Elements<T>().ToList())
        {
            if (TypeOf(existing) == type)
            {
                existing.Remove();
            }
        }

        var reference = new T();
        switch (reference)
        {
            case W.HeaderReference header:
                header.Type = type;
                header.Id = relationshipId;
                break;
            case W.FooterReference footer:
                footer.Type = type;
                footer.Id = relationshipId;
                break;
        }

        element.PrependChild(reference);
    }

    static W.HeaderFooterValues? TypeOf(OpenXmlElement element) =>
        element switch
        {
            W.HeaderReference header => header.Type?.Value,
            W.FooterReference footer => footer.Type?.Value,
            _ => null
        };

    Document RequireDocument()
    {
        if (document == null)
        {
            throw new InvalidOperationException(
                "Headers and footers live in their own package parts, so they need a document to be added to. This section is not attached to one.");
        }

        return document;
    }

    /// <summary>
    /// The underlying OpenXML element, with any pending page setup applied.
    /// </summary>
    public W.SectionProperties ToOpenXml()
    {
        Flush();
        return element;
    }

    internal void Flush() =>
        pageSetup?.ApplyTo(element);
}
