namespace OpenXmlKit.Word;

public partial class Font
{
    /// <summary>
    /// Overlays another font on this one: everything it states wins, everything it leaves unstated
    /// is left alone.
    /// </summary>
    /// <remarks>
    /// This is the whole of how Word resolves character formatting — the document defaults are laid
    /// down first, then each style in turn, then the direct formatting, each overlaying only what
    /// it actually says. Modelling "unstated" as a distinct value rather than as a default is what
    /// makes that expressible; a plain <c>bool</c> for bold would overwrite an inherited on with
    /// its own false at every step.
    /// </remarks>
    public void MergeFrom(Font higher)
    {
        NameAscii = higher.NameAscii ?? NameAscii;
        NameHighAnsi = higher.NameHighAnsi ?? NameHighAnsi;
        NameComplexScript = higher.NameComplexScript ?? NameComplexScript;
        NameEastAsia = higher.NameEastAsia ?? NameEastAsia;
        Size = higher.Size ?? Size;
        SizeComplexScript = higher.SizeComplexScript ?? SizeComplexScript;

        Bold = Overlay(Bold, higher.Bold);
        BoldComplexScript = Overlay(BoldComplexScript, higher.BoldComplexScript);
        Italic = Overlay(Italic, higher.Italic);
        ItalicComplexScript = Overlay(ItalicComplexScript, higher.ItalicComplexScript);
        Strike = Overlay(Strike, higher.Strike);
        DoubleStrike = Overlay(DoubleStrike, higher.DoubleStrike);
        SmallCaps = Overlay(SmallCaps, higher.SmallCaps);
        AllCaps = Overlay(AllCaps, higher.AllCaps);
        Outline = Overlay(Outline, higher.Outline);
        Shadow = Overlay(Shadow, higher.Shadow);
        Emboss = Overlay(Emboss, higher.Emboss);
        Imprint = Overlay(Imprint, higher.Imprint);
        Hidden = Overlay(Hidden, higher.Hidden);
        NoProof = Overlay(NoProof, higher.NoProof);
        RightToLeft = Overlay(RightToLeft, higher.RightToLeft);

        Underline = higher.Underline ?? Underline;
        UnderlineColor = higher.UnderlineColor ?? UnderlineColor;
        Color = higher.Color ?? Color;
        Highlight = higher.Highlight ?? Highlight;
        VerticalPosition = higher.VerticalPosition ?? VerticalPosition;
        CharacterSpacing = higher.CharacterSpacing ?? CharacterSpacing;
        Scale = higher.Scale ?? Scale;
        Position = higher.Position ?? Position;
        Language = higher.Language ?? Language;
        LanguageEastAsia = higher.LanguageEastAsia ?? LanguageEastAsia;
        LanguageComplexScript = higher.LanguageComplexScript ?? LanguageComplexScript;

        if (!higher.Shading.IsEmpty)
        {
            Shading.CopyFrom(higher.Shading);
        }

        if (!higher.Border.IsEmpty)
        {
            Border.CopyFrom(higher.Border);
        }

        // Not the style id: what is being resolved is the formatting the style contributed, and
        // carrying the name of the style down would make the result look like a reference to it.
    }

    static Toggle Overlay(Toggle lower, Toggle higher) =>
        higher.IsSet ? higher : lower;
}
