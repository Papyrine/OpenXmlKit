namespace OpenXmlKit.Word;

/// <summary>
/// A list, as referenced by the paragraphs belonging to it.
/// </summary>
/// <remarks>
/// Two things stand behind a list in the file format: an abstract definition describing what each
/// of the nine levels looks like, and an instance pointing at it. Paragraphs reference the
/// instance, which is why two lists can share one appearance and still number independently — see
/// <see cref="Numbering.Restart"/>.
/// </remarks>
public class ListDefinition
{
    readonly List<ListLevel> levels;

    internal ListDefinition(int abstractId, int numberingId, List<ListLevel> levels)
    {
        AbstractId = abstractId;
        NumberingId = numberingId;
        this.levels = levels;
    }

    /// <summary>
    /// The id paragraphs reference.
    /// </summary>
    public int NumberingId { get; }

    internal int AbstractId { get; }

    /// <summary>
    /// The nine levels, outermost first.
    /// </summary>
    public IReadOnlyList<ListLevel> Levels => levels;

    /// <summary>
    /// One level, by depth.
    /// </summary>
    public ListLevel this[int depth] => levels[depth];

    /// <summary>
    /// Membership of this list at the given depth, for assigning to a paragraph's format.
    /// </summary>
    public ListMembership At(int depth = 0) =>
        new(NumberingId, depth);

    /// <summary>
    /// Configures one level.
    /// </summary>
    public ListDefinition Level(int depth, Action<ListLevel> configure)
    {
        configure(levels[depth]);
        return this;
    }

    /// <summary>
    /// Configures every level.
    /// </summary>
    public ListDefinition EachLevel(Action<ListLevel> configure)
    {
        foreach (var level in levels)
        {
            configure(level);
        }

        return this;
    }
}
