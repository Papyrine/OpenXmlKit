namespace OpenXmlKit.Word;

/// <summary>
/// Page geometry for a section: paper, orientation, margins, and columns.
/// </summary>
public class PageSetup :
    IPageSetupView
{
    public Length? PageWidth { get; set; }
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

    public Length? LeftMargin { get; set; }
    public Length? RightMargin { get; set; }
    public Length? TopMargin { get; set; }
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

    public SectionStart? Start { get; set; }

    /// <summary>
    /// How many columns the text flows into.
    /// </summary>
    public int? ColumnCount { get; set; }

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

    public void SetMargins(Length left, Length top, Length right, Length bottom)
    {
        LeftMargin = left;
        TopMargin = top;
        RightMargin = right;
        BottomMargin = bottom;
    }

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

    public PageSetup Clone()
    {
        var clone = new PageSetup();
        clone.CopyFrom(this);
        return clone;
    }

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
            properties.GetFirstChild<W.SectionType>()?.Remove();
            properties.AppendChild(
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

    static void Replace<T>(W.SectionProperties properties, T element)
        where T : OpenXmlElement
    {
        properties.GetFirstChild<T>()?.Remove();
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
