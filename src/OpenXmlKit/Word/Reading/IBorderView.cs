namespace OpenXmlKit.Word;

/// <summary>
/// One edge of a border box, as read.
/// </summary>
public interface IBorderView
{
    BorderStyle? Style { get; }
    Length? Width { get; }
    Color? Color { get; }
    Length? Space { get; }
    bool Shadow { get; }
    bool IsEmpty { get; }
}
