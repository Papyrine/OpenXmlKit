namespace OpenXmlKit.Word;

/// <summary>
/// One level of a list definition, as read.
/// </summary>
public interface IListLevelView
{
    int Depth { get; }
    NumberFormat Format { get; }

    /// <summary>
    /// The marker, with <c>%n</c> standing for the number at level n.
    /// </summary>
    string Text { get; }

    int StartAt { get; }
    ListLevelAlignment Alignment { get; }
    Length Indent { get; }
    Length Hanging { get; }
    ListTrailingCharacter TrailingCharacter { get; }
    IFontView Font { get; }
    int? RestartAfterLevel { get; }
}
