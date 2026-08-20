namespace OpenXmlKit.Word;

/// <summary>
/// Page geometry, as read.
/// </summary>
public interface IPageSetupView
{
    /// <inheritdoc cref="PageSetup.PageWidth"/>
    Length? PageWidth { get; }

    /// <inheritdoc cref="PageSetup.PageHeight"/>
    Length? PageHeight { get; }

    /// <inheritdoc cref="PageSetup.Orientation"/>
    PageOrientation? Orientation { get; }

    /// <inheritdoc cref="PageSetup.LeftMargin"/>
    Length? LeftMargin { get; }

    /// <inheritdoc cref="PageSetup.RightMargin"/>
    Length? RightMargin { get; }

    /// <inheritdoc cref="PageSetup.TopMargin"/>
    Length? TopMargin { get; }

    /// <inheritdoc cref="PageSetup.BottomMargin"/>
    Length? BottomMargin { get; }

    /// <inheritdoc cref="PageSetup.HeaderDistance"/>
    Length? HeaderDistance { get; }

    /// <inheritdoc cref="PageSetup.FooterDistance"/>
    Length? FooterDistance { get; }

    /// <inheritdoc cref="PageSetup.Gutter"/>
    Length? Gutter { get; }

    /// <inheritdoc cref="PageSetup.Start"/>
    SectionStart? Start { get; }

    /// <inheritdoc cref="PageSetup.ColumnCount"/>
    int? ColumnCount { get; }

    /// <inheritdoc cref="PageSetup.ColumnSpacing"/>
    Length? ColumnSpacing { get; }

    /// <inheritdoc cref="PageSetup.ColumnSeparator"/>
    bool ColumnSeparator { get; }

    /// <inheritdoc cref="PageSetup.DifferentFirstPage"/>
    bool DifferentFirstPage { get; }

    /// <inheritdoc cref="PageSetup.DifferentOddAndEvenPages"/>
    bool DifferentOddAndEvenPages { get; }

    /// <inheritdoc cref="PageSetup.PageNumberStart"/>
    int? PageNumberStart { get; }
}
