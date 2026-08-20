namespace OpenXmlKit.Word;

/// <summary>
/// Character formatting, as read.
/// </summary>
/// <remarks>
/// The same properties <see cref="Font"/> carries, without the setters. Reading a document should
/// not hand back something that looks like it can be written to and then quietly is not.
/// </remarks>
public interface IFontView
{
    /// <inheritdoc cref="Font.Name"/>
    string? Name { get; }

    /// <inheritdoc cref="Font.NameAscii"/>
    string? NameAscii { get; }

    /// <inheritdoc cref="Font.NameHighAnsi"/>
    string? NameHighAnsi { get; }

    /// <inheritdoc cref="Font.NameComplexScript"/>
    string? NameComplexScript { get; }

    /// <inheritdoc cref="Font.NameEastAsia"/>
    string? NameEastAsia { get; }

    /// <inheritdoc cref="Font.Size"/>
    Length? Size { get; }

    /// <inheritdoc cref="Font.SizeComplexScript"/>
    Length? SizeComplexScript { get; }

    /// <inheritdoc cref="Font.Bold"/>
    Toggle Bold { get; }

    /// <inheritdoc cref="Font.BoldComplexScript"/>
    Toggle BoldComplexScript { get; }

    /// <inheritdoc cref="Font.Italic"/>
    Toggle Italic { get; }

    /// <inheritdoc cref="Font.ItalicComplexScript"/>
    Toggle ItalicComplexScript { get; }

    /// <inheritdoc cref="Font.Underline"/>
    UnderlineStyle? Underline { get; }

    /// <inheritdoc cref="Font.UnderlineColor"/>
    Color? UnderlineColor { get; }

    /// <inheritdoc cref="Font.Strike"/>
    Toggle Strike { get; }

    /// <inheritdoc cref="Font.DoubleStrike"/>
    Toggle DoubleStrike { get; }

    /// <inheritdoc cref="Font.SmallCaps"/>
    Toggle SmallCaps { get; }

    /// <inheritdoc cref="Font.AllCaps"/>
    Toggle AllCaps { get; }

    /// <inheritdoc cref="Font.Outline"/>
    Toggle Outline { get; }

    /// <inheritdoc cref="Font.Shadow"/>
    Toggle Shadow { get; }

    /// <inheritdoc cref="Font.Emboss"/>
    Toggle Emboss { get; }

    /// <inheritdoc cref="Font.Imprint"/>
    Toggle Imprint { get; }

    /// <inheritdoc cref="Font.Hidden"/>
    Toggle Hidden { get; }

    /// <inheritdoc cref="Font.NoProof"/>
    Toggle NoProof { get; }

    /// <inheritdoc cref="Font.RightToLeft"/>
    Toggle RightToLeft { get; }

    /// <inheritdoc cref="Font.Color"/>
    Color? Color { get; }

    /// <inheritdoc cref="Font.Highlight"/>
    HighlightColor? Highlight { get; }

    /// <inheritdoc cref="Font.Shading"/>
    IShadingView Shading { get; }

    /// <inheritdoc cref="Font.Border"/>
    IBorderView Border { get; }

    /// <inheritdoc cref="Font.VerticalPosition"/>
    VerticalTextPosition? VerticalPosition { get; }

    /// <inheritdoc cref="Font.CharacterSpacing"/>
    Length? CharacterSpacing { get; }

    /// <inheritdoc cref="Font.Scale"/>
    int? Scale { get; }

    /// <inheritdoc cref="Font.Position"/>
    Length? Position { get; }

    /// <inheritdoc cref="Font.Language"/>
    string? Language { get; }

    /// <inheritdoc cref="Font.LanguageEastAsia"/>
    string? LanguageEastAsia { get; }

    /// <inheritdoc cref="Font.LanguageComplexScript"/>
    string? LanguageComplexScript { get; }

    /// <summary>
    /// The character style named on the run, if any. What that style resolves to is
    /// <see cref="FormattingResolver"/>'s business.
    /// </summary>
    string? StyleId { get; }

    /// <inheritdoc cref="Font.IsEmpty"/>
    bool IsEmpty { get; }
}
