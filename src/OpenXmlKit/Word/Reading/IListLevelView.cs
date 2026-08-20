namespace OpenXmlKit.Word;

/// <summary>
/// One level of a list definition, as read.
/// </summary>
public interface IListLevelView
{
    /// <inheritdoc cref="ListLevel.Depth"/>
    int Depth { get; }

    /// <inheritdoc cref="ListLevel.Format"/>
    NumberFormat Format { get; }

    /// <summary>
    /// The marker, with <c>%n</c> standing for the number at level n.
    /// </summary>
    string Text { get; }

    /// <inheritdoc cref="ListLevel.StartAt"/>
    int StartAt { get; }

    /// <inheritdoc cref="ListLevel.Alignment"/>
    ListLevelAlignment Alignment { get; }

    /// <inheritdoc cref="ListLevel.Indent"/>
    Length Indent { get; }

    /// <inheritdoc cref="ListLevel.Hanging"/>
    Length Hanging { get; }

    /// <inheritdoc cref="ListLevel.TrailingCharacter"/>
    ListTrailingCharacter TrailingCharacter { get; }

    /// <inheritdoc cref="ListLevel.Font"/>
    IFontView Font { get; }

    /// <inheritdoc cref="ListLevel.RestartAfterLevel"/>
    int? RestartAfterLevel { get; }
}
