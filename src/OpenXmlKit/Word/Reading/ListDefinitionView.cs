namespace OpenXmlKit.Word;

/// <summary>
/// A list definition in a document being read.
/// </summary>
public readonly struct ListDefinitionView
{
    readonly W.AbstractNum abstractNum;

    internal ListDefinitionView(int numberingId, W.AbstractNum abstractNum)
    {
        NumberingId = numberingId;
        this.abstractNum = abstractNum;
    }

    /// <summary>
    /// The id paragraphs reference, as carried by <see cref="ListMembership"/>.
    /// </summary>
    public int NumberingId { get; }

    /// <summary>
    /// The nine levels, outermost first.
    /// </summary>
    public IEnumerable<IListLevelView> Levels
    {
        get
        {
            foreach (var level in abstractNum.Elements<W.Level>())
            {
                yield return ListLevel.Read(level);
            }
        }
    }

    /// <summary>
    /// One level, or null when the definition does not go that deep.
    /// </summary>
    public IListLevelView? this[int depth]
    {
        get
        {
            foreach (var level in abstractNum.Elements<W.Level>())
            {
                if (level.LevelIndex?.Value == depth)
                {
                    return ListLevel.Read(level);
                }
            }

            return null;
        }
    }

    /// <summary>
    /// The underlying OpenXML element, for anything this view does not expose.
    /// </summary>
    public W.AbstractNum ToOpenXml() =>
        abstractNum;
}
