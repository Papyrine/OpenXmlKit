namespace OpenXmlKit.Word;

/// <summary>
/// Page geometry for a section: paper, orientation, margins, and columns.
/// </summary>
public class PageSetup :
    IPageSetupView
{
    /// <summary>
    /// The width of the page.
    /// </summary>
    public Length? PageWidth { get; set; }

    /// <summary>
    /// The height of the page.
    /// </summary>
    public Length? PageHeight { get; set; }

    /// <summary>
    /// Portrait or landscape.
    /// </summary>
    /// <remarks>
    /// Setting this does not swap <see cref="PageWidth"/> and <see cref="PageHeight"/>, and Word
    /// takes the measurements literally — so landscape with portrait dimensions renders portrait.
    /// <see cref="SetPaper"/> keeps the two consistent.
    /// </remarks>
    public PageOrientation? Orientation { get; set; }

    /// <summary>
    /// Space between the left edge of the page and the text.
    /// </summary>
    public Length? LeftMargin { get; set; }

    /// <summary>
    /// Space between the text and the right edge.
    /// </summary>
    public Length? RightMargin { get; set; }

    /// <summary>
    /// Space between the top edge of the page and the text.
    /// </summary>
    public Length? TopMargin { get; set; }

    /// <summary>
    /// Space between the text and the bottom edge.
    /// </summary>
    public Length? BottomMargin { get; set; }

    /// <summary>
    /// Distance from the top of the page to the header.
    /// </summary>
    public Length? HeaderDistance { get; set; }

    /// <summary>
    /// Distance from the bottom of the page to the footer.
    /// </summary>
    public Length? FooterDistance { get; set; }

    /// <summary>
    /// Extra margin on the binding edge.
    /// </summary>
    public Length? Gutter { get; set; }

    /// <summary>
    /// Where the content after this section break begins.
    /// </summary>
    public SectionStart? Start { get; set; }

    /// <summary>
    /// How many columns the text flows into.
    /// </summary>
    public int? ColumnCount { get; set; }

    /// <summary>
    /// The gutter between columns, when there is more than one.
    /// </summary>
    public Length? ColumnSpacing { get; set; }

    /// <summary>
    /// Draws a vertical rule between columns.
    /// </summary>
    public bool ColumnSeparator { get; set; }

    /// <summary>
    /// Lets the first page of the section carry its own header and footer.
    /// </summary>
    /// <remarks>
    /// A <see cref="HeaderFooterKind.First"/> header added without this is written to the document
    /// and never shown.
    /// </remarks>
    public bool DifferentFirstPage { get; set; }

    /// <summary>
    /// Lets left and right pages carry different headers and footers.
    /// </summary>
    /// <remarks>
    /// Word stores the odd/even switch once for the whole document rather than per section, so
    /// this is written to the document settings — see <c>Document.Settings</c>.
    /// </remarks>
    public bool DifferentOddAndEvenPages { get; set; }

    /// <summary>
    /// The number the section's page numbering starts from.
    /// </summary>
    public int? PageNumberStart { get; set; }

    /// <summary>
    /// Sets width and height together, and the orientation to match.
    /// </summary>
    public void SetPaper(Length width, Length height, PageOrientation orientation = PageOrientation.Portrait)
    {
        // Word stores the measurements in the orientation they are used, so landscape is a wide
        // page rather than a tall page with a flag on it.
        var isLandscape = orientation == PageOrientation.Landscape;
        PageWidth = isLandscape ? height : width;
        PageHeight = isLandscape ? width : height;
        Orientation = orientation;
    }

    /// <summary>
    /// A4, in the given orientation.
    /// </summary>
    public void SetA4(PageOrientation orientation = PageOrientation.Portrait) =>
        SetPaper(Length.FromMillimeters(210), Length.FromMillimeters(297), orientation);

    /// <summary>
    /// US Letter, in the given orientation.
    /// </summary>
    public void SetLetter(PageOrientation orientation = PageOrientation.Portrait) =>
        SetPaper(Length.FromInches(8.5), Length.FromInches(11), orientation);

    /// <summary>
    /// Sets all four page margins at once.
    /// </summary>
    public void SetMargins(Length all) =>
        SetMargins(all, all, all, all);

    /// <summary>
    /// Sets all four page margins at once.
    /// </summary>
    public void SetMargins(Length left, Length top, Length right, Length bottom)
    {
        LeftMargin = left;
        TopMargin = top;
        RightMargin = right;
        BottomMargin = bottom;
    }

    /// <summary>
    /// Whether anything at all is stated. An empty format writes no properties element,
    /// leaving the style hierarchy to resolve every value.
    /// </summary>
    public bool IsEmpty =>
        PageWidth == null &&
        PageHeight == null &&
        Orientation == null &&
        LeftMargin == null &&
        RightMargin == null &&
        TopMargin == null &&
        BottomMargin == null &&
        HeaderDistance == null &&
        FooterDistance == null &&
        Gutter == null &&
        Start == null &&
        ColumnCount == null &&
        ColumnSpacing == null &&
        !ColumnSeparator &&
        !DifferentFirstPage &&
        !DifferentOddAndEvenPages &&
        PageNumberStart == null;

    /// <summary>
    /// An independent copy, so the two can diverge.
    /// </summary>
    public PageSetup Clone()
    {
        var clone = new PageSetup();
        clone.CopyFrom(this);
        return clone;
    }

    /// <summary>
    /// Overwrites every property with the other value, stated or not.
    /// </summary>
    public void CopyFrom(PageSetup other)
    {
        PageWidth = other.PageWidth;
        PageHeight = other.PageHeight;
        Orientation = other.Orientation;
        LeftMargin = other.LeftMargin;
        RightMargin = other.RightMargin;
        TopMargin = other.TopMargin;
        BottomMargin = other.BottomMargin;
        HeaderDistance = other.HeaderDistance;
        FooterDistance = other.FooterDistance;
        Gutter = other.Gutter;
        Start = other.Start;
        ColumnCount = other.ColumnCount;
        ColumnSpacing = other.ColumnSpacing;
        ColumnSeparator = other.ColumnSeparator;
        DifferentFirstPage = other.DifferentFirstPage;
        DifferentOddAndEvenPages = other.DifferentOddAndEvenPages;
        PageNumberStart = other.PageNumberStart;
    }

    internal void ApplyTo(W.SectionProperties properties)
    {
        if (Start is { } start)
        {
            Replace(
                properties,
                new W.SectionType
                {
                    Val = start.ToOpenXml()
                });
        }

        if (PageWidth != null ||
            PageHeight != null ||
            Orientation != null)
        {
            var size = new W.PageSize();
            if (PageWidth is { } width)
            {
                size.Width = (uint) width.Twips;
            }

            if (PageHeight is { } height)
            {
                size.Height = (uint) height.Twips;
            }

            if (Orientation is { } orientation)
            {
                size.Orient = orientation.ToOpenXml();
            }

            Replace(properties, size);
        }

        if (LeftMargin != null ||
            RightMargin != null ||
            TopMargin != null ||
            BottomMargin != null ||
            HeaderDistance != null ||
            FooterDistance != null ||
            Gutter != null)
        {
            // Every attribute of pgMar is required, so the ones the caller did not state fall back
            // to what Word uses for a new document rather than to zero.
            var margin = new W.PageMargin
            {
                Left = (uint) (LeftMargin ?? Length.FromInches(1)).Twips,
                Right = (uint) (RightMargin ?? Length.FromInches(1)).Twips,
                Top = (TopMargin ?? Length.FromInches(1)).Twips,
                Bottom = (BottomMargin ?? Length.FromInches(1)).Twips,
                Header = (uint) (HeaderDistance ?? Length.FromInches(0.5)).Twips,
                Footer = (uint) (FooterDistance ?? Length.FromInches(0.5)).Twips,
                Gutter = (uint) (Gutter ?? Length.Zero).Twips
            };
            Replace(properties, margin);
        }

        if (ColumnCount is { } columns)
        {
            var element = new W.Columns
            {
                ColumnCount = (short) columns,
                EqualWidth = true,
                Separator = ColumnSeparator
            };
            if (ColumnSpacing is { } spacing)
            {
                element.Space = spacing.Twips.ToString(CultureInfo.InvariantCulture);
            }

            Replace(properties, element);
        }

        if (PageNumberStart is { } pageNumberStart)
        {
            Replace(
                properties,
                new W.PageNumberType
                {
                    Start = pageNumberStart
                });
        }

        if (DifferentFirstPage)
        {
            Replace(properties, new W.TitlePage());
        }
    }

    // The CT_SectPr sequence. Stated here rather than assigned through a typed property, which is
    // what every other properties container in this library uses, because sectPr has none: header
    // and footer references repeat, so the SDK generates no typed children for it at all.
    static readonly string[] order =
    [
        "headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz", "pgMar",
        "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt", "vAlign",
        "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings",
        "footnoteColumns", "sectPrChange"
    ];

    /// <remarks>
    /// Appending would be enough only while this library owns the whole of sectPr. The moment a
    /// child it does not model is there — a pgBorders put in through the escape hatch, or anything
    /// already in an opened template — an appended pgSz lands after it, and a sectPr out of
    /// sequence is a document Word offers to repair.
    /// </remarks>
    static void Replace<T>(W.SectionProperties properties, T element)
        where T : OpenXmlElement
    {
        properties.GetFirstChild<T>()?.Remove();

        var position = Array.IndexOf(order, element.LocalName);
        foreach (var child in properties.ChildElements)
        {
            // Children this does not know about are stepped over rather than displaced: an
            // extension element has no place in the sequence to be measured against.
            var childPosition = Array.IndexOf(order, child.LocalName);
            if (childPosition >= 0 &&
                childPosition > position)
            {
                properties.InsertBefore(element, child);
                return;
            }
        }

        properties.AppendChild(element);
    }

    internal void ReadFrom(W.SectionProperties properties)
    {
        if (properties.GetFirstChild<W.PageSize>() is { } size)
        {
            if (size.Width is { HasValue: true } width)
            {
                PageWidth = Length.FromTwips(width.Value);
            }

            if (size.Height is { HasValue: true } height)
            {
                PageHeight = Length.FromTwips(height.Value);
            }

            if (size.Orient is { HasValue: true } orient)
            {
                Orientation = Map.ToOrientation(orient.Value);
            }
        }

        if (properties.GetFirstChild<W.PageMargin>() is { } margin)
        {
            if (margin.Left is { HasValue: true } left)
            {
                LeftMargin = Length.FromTwips(left.Value);
            }

            if (margin.Right is { HasValue: true } right)
            {
                RightMargin = Length.FromTwips(right.Value);
            }

            if (margin.Top is { HasValue: true } top)
            {
                TopMargin = Length.FromTwips(top.Value);
            }

            if (margin.Bottom is { HasValue: true } bottom)
            {
                BottomMargin = Length.FromTwips(bottom.Value);
            }

            if (margin.Header is { HasValue: true } header)
            {
                HeaderDistance = Length.FromTwips(header.Value);
            }

            if (margin.Footer is { HasValue: true } footer)
            {
                FooterDistance = Length.FromTwips(footer.Value);
            }

            if (margin.Gutter is { HasValue: true } gutter)
            {
                Gutter = Length.FromTwips(gutter.Value);
            }
        }

        if (properties.GetFirstChild<W.SectionType>()?.Val is { HasValue: true } start)
        {
            Start = Map.ToSectionStart(start.Value);
        }

        if (properties.GetFirstChild<W.Columns>() is { } columns)
        {
            if (columns.ColumnCount is { HasValue: true } count)
            {
                ColumnCount = count.Value;
            }

            ColumnSeparator = columns.Separator?.Value == true;

            if (columns.Space?.Value is { } space &&
                double.TryParse(space, NumberStyles.Float, CultureInfo.InvariantCulture, out var spacing))
            {
                ColumnSpacing = Length.FromTwips(spacing);
            }
        }

        if (properties.GetFirstChild<W.PageNumberType>()?.Start is { HasValue: true } pageNumberStart)
        {
            PageNumberStart = pageNumberStart.Value;
        }

        DifferentFirstPage = properties.GetFirstChild<W.TitlePage>() != null;
    }
}
