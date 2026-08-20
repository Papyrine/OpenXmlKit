namespace OpenXmlKit.Word;

/// <summary>
/// A background fill, as read.
/// </summary>
public interface IShadingView
{
    /// <inheritdoc cref="Shading.BackgroundColor"/>
    Color? BackgroundColor { get; }

    /// <inheritdoc cref="Shading.PatternColor"/>
    Color? PatternColor { get; }

    /// <inheritdoc cref="Shading.Pattern"/>
    ShadingPattern Pattern { get; }

    /// <inheritdoc cref="Shading.IsEmpty"/>
    bool IsEmpty { get; }
}
