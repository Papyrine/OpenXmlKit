namespace OpenXmlKit.Word;

/// <summary>
/// One edge of a border box, as read.
/// </summary>
public interface IBorderView
{
    /// <inheritdoc cref="Border.Style"/>
    BorderStyle? Style { get; }

    /// <inheritdoc cref="Border.Width"/>
    Length? Width { get; }

    /// <inheritdoc cref="Border.Color"/>
    Color? Color { get; }

    /// <inheritdoc cref="Border.Space"/>
    Length? Space { get; }

    /// <inheritdoc cref="Border.Shadow"/>
    bool Shadow { get; }

    /// <inheritdoc cref="Border.IsEmpty"/>
    bool IsEmpty { get; }
}
