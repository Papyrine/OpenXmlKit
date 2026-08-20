namespace OpenXmlKit.Word;

/// <summary>
/// The six edges of a border box, as read.
/// </summary>
public interface IBordersView
{
    IBorderView Top { get; }
    IBorderView Bottom { get; }
    IBorderView Left { get; }
    IBorderView Right { get; }
    IBorderView InsideHorizontal { get; }
    IBorderView InsideVertical { get; }
    bool IsEmpty { get; }
}
