namespace OpenXmlKit.Word;

public partial class Font
{
    // Explicit, because C# does not accept a covariant return type for an implicit interface
    // implementation: Font.Shading returns the concrete Shading, which is what the write side
    // wants, and that does not satisfy IFontView.Shading on its own.
    IShadingView IFontView.Shading => Shading;
    IBorderView IFontView.Border => Border;
}
