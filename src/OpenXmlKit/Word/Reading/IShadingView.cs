namespace OpenXmlKit.Word;

/// <summary>
/// A background fill, as read.
/// </summary>
public interface IShadingView
{
    Color? BackgroundColor { get; }
    Color? PatternColor { get; }
    ShadingPattern Pattern { get; }
    bool IsEmpty { get; }
}
