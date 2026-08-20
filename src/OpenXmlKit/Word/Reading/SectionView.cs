namespace OpenXmlKit.Word;

/// <summary>
/// A section in a document being read: a run of pages sharing one page setup and one set of
/// headers and footers.
/// </summary>
public readonly struct SectionView
{
    readonly W.SectionProperties element;
    readonly MainDocumentPart main;

    internal SectionView(W.SectionProperties element, MainDocumentPart main)
    {
        this.element = element;
        this.main = main;
    }

    /// <summary>
    /// The page size, margins and columns the section states.
    /// </summary>
    public IPageSetupView PageSetup
    {
        get
        {
            var setup = new PageSetup();
            setup.ReadFrom(element);
            return setup;
        }
    }

    /// <summary>
    /// The headers this section declares, by which pages they apply to.
    /// </summary>
    public IEnumerable<KeyValuePair<HeaderFooterKind, BlockContainerView>> Headers
    {
        get
        {
            foreach (var reference in element.Elements<W.HeaderReference>())
            {
                if (Part<HeaderPart>(reference.Id?.Value)?.Header is { } header)
                {
                    yield return new(Kind(reference.Type?.Value), new(header));
                }
            }
        }
    }

    /// <summary>
    /// The footers this section declares.
    /// </summary>
    public IEnumerable<KeyValuePair<HeaderFooterKind, BlockContainerView>> Footers
    {
        get
        {
            foreach (var reference in element.Elements<W.FooterReference>())
            {
                if (Part<FooterPart>(reference.Id?.Value)?.Footer is { } footer)
                {
                    yield return new(Kind(reference.Type?.Value), new(footer));
                }
            }
        }
    }

    static HeaderFooterKind Kind(W.HeaderFooterValues? type) =>
        type == null ? HeaderFooterKind.Default : Map.ToHeaderFooterKind(type.Value);

    T? Part<T>(string? relationshipId)
        where T : OpenXmlPart
    {
        if (relationshipId == null)
        {
            return null;
        }

        return main.GetPartById(relationshipId) as T;
    }

    /// <summary>
    /// The underlying OpenXML element, for anything this view does not expose.
    /// </summary>
    public W.SectionProperties ToOpenXml() =>
        element;
}
