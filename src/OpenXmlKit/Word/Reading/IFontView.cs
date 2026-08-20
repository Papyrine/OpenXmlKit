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
    string? Name { get; }
    string? NameAscii { get; }
    string? NameHighAnsi { get; }
    string? NameComplexScript { get; }
    string? NameEastAsia { get; }
    Length? Size { get; }
    Length? SizeComplexScript { get; }
    Toggle Bold { get; }
    Toggle BoldComplexScript { get; }
    Toggle Italic { get; }
    Toggle ItalicComplexScript { get; }
    UnderlineStyle? Underline { get; }
    Color? UnderlineColor { get; }
    Toggle Strike { get; }
    Toggle DoubleStrike { get; }
    Toggle SmallCaps { get; }
    Toggle AllCaps { get; }
    Toggle Outline { get; }
    Toggle Shadow { get; }
    Toggle Emboss { get; }
    Toggle Imprint { get; }
    Toggle Hidden { get; }
    Toggle NoProof { get; }
    Toggle RightToLeft { get; }
    Color? Color { get; }
    HighlightColor? Highlight { get; }
    IShadingView Shading { get; }
    IBorderView Border { get; }
    VerticalTextPosition? VerticalPosition { get; }
    Length? CharacterSpacing { get; }
    int? Scale { get; }
    Length? Position { get; }
    string? Language { get; }
    string? LanguageEastAsia { get; }
    string? LanguageComplexScript { get; }

    /// <summary>
    /// The character style named on the run, if any. What that style resolves to is
    /// <see cref="FormattingResolver"/>'s business.
    /// </summary>
    string? StyleId { get; }

    bool IsEmpty { get; }
}
