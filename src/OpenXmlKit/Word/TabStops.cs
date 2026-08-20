namespace OpenXmlKit.Word;

/// <summary>
/// A single tab stop.
/// </summary>
/// <param name="Position">
/// Distance from the left margin.
/// </param>
/// <param name="Alignment">
/// How text sits against the stop.
/// </param>
/// <param name="Leader">
/// What fills the gap the tab jumps across.
/// </param>
public readonly record struct TabStop(
    Length Position,
    TabAlignment Alignment = TabAlignment.Left,
    TabLeader Leader = TabLeader.None);

/// <summary>
/// The tab stops a paragraph declares, kept in position order.
/// </summary>
public class TabStops :
    IEnumerable<TabStop>
{
    readonly List<TabStop> stops = [];

    public int Count => stops.Count;

    public TabStop this[int index] => stops[index];

    public bool IsEmpty => stops.Count == 0;

    /// <summary>
    /// Adds a stop, keeping the list in position order. A stop at a position already taken
    /// replaces the one there — which is how Word itself behaves, since a position identifies a
    /// stop.
    /// </summary>
    public TabStops Add(Length position, TabAlignment alignment = TabAlignment.Left, TabLeader leader = TabLeader.None) =>
        Add(new(position, alignment, leader));

    public TabStops Add(TabStop stop)
    {
        var index = stops.FindIndex(_ => _.Position == stop.Position);
        if (index >= 0)
        {
            stops[index] = stop;
            return this;
        }

        index = stops.FindIndex(_ => _.Position > stop.Position);
        if (index < 0)
        {
            stops.Add(stop);
        }
        else
        {
            stops.Insert(index, stop);
        }

        return this;
    }

    /// <summary>
    /// Cancels a stop inherited from the style. Removing one that was declared here is
    /// <see cref="Remove"/>; this states that a stop the style declares should not apply.
    /// </summary>
    public TabStops Clear(Length position) =>
        Add(new(position, TabAlignment.Clear));

    public bool Remove(Length position) =>
        stops.RemoveAll(_ => _.Position == position) > 0;

    public void RemoveAll() =>
        stops.Clear();

    public void CopyFrom(TabStops other)
    {
        stops.Clear();
        stops.AddRange(other.stops);
    }

    public IEnumerator<TabStop> GetEnumerator() =>
        stops.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();

    internal W.Tabs? ToOpenXml()
    {
        if (IsEmpty)
        {
            return null;
        }

        var tabs = new W.Tabs();
        foreach (var stop in stops)
        {
            tabs.Append(
                new W.TabStop
                {
                    Val = ToOpenXml(stop.Alignment),
                    Position = stop.Position.Twips,
                    Leader = ToOpenXml(stop.Leader)
                });
        }

        return tabs;
    }

    internal void ReadFrom(W.Tabs? tabs)
    {
        stops.Clear();
        if (tabs == null)
        {
            return;
        }

        foreach (var tab in tabs.Elements<W.TabStop>())
        {
            if (tab.Position is not { HasValue: true } position)
            {
                continue;
            }

            stops.Add(
                new(
                    Length.FromTwips(position.Value),
                    tab.Val is { HasValue: true } alignment ? ToAlignment(alignment.Value) : TabAlignment.Left,
                    tab.Leader is { HasValue: true } leader ? ToLeader(leader.Value) : TabLeader.None));
        }
    }

    static W.TabStopValues ToOpenXml(TabAlignment alignment) =>
        alignment switch
        {
            TabAlignment.Center => W.TabStopValues.Center,
            TabAlignment.Right => W.TabStopValues.Right,
            TabAlignment.Decimal => W.TabStopValues.Decimal,
            TabAlignment.Bar => W.TabStopValues.Bar,
            TabAlignment.Clear => W.TabStopValues.Clear,
            _ => W.TabStopValues.Left
        };

    static TabAlignment ToAlignment(W.TabStopValues value)
    {
        if (value == W.TabStopValues.Center)
        {
            return TabAlignment.Center;
        }

        if (value == W.TabStopValues.Right)
        {
            return TabAlignment.Right;
        }

        if (value == W.TabStopValues.Decimal)
        {
            return TabAlignment.Decimal;
        }

        if (value == W.TabStopValues.Bar)
        {
            return TabAlignment.Bar;
        }

        if (value == W.TabStopValues.Clear)
        {
            return TabAlignment.Clear;
        }

        return TabAlignment.Left;
    }

    static W.TabStopLeaderCharValues ToOpenXml(TabLeader leader) =>
        leader switch
        {
            TabLeader.Dots => W.TabStopLeaderCharValues.Dot,
            TabLeader.Dashes => W.TabStopLeaderCharValues.Hyphen,
            TabLeader.Underscore => W.TabStopLeaderCharValues.Underscore,
            TabLeader.Heavy => W.TabStopLeaderCharValues.Heavy,
            TabLeader.MiddleDot => W.TabStopLeaderCharValues.MiddleDot,
            _ => W.TabStopLeaderCharValues.None
        };

    static TabLeader ToLeader(W.TabStopLeaderCharValues value)
    {
        if (value == W.TabStopLeaderCharValues.Dot)
        {
            return TabLeader.Dots;
        }

        if (value == W.TabStopLeaderCharValues.Hyphen)
        {
            return TabLeader.Dashes;
        }

        if (value == W.TabStopLeaderCharValues.Underscore)
        {
            return TabLeader.Underscore;
        }

        if (value == W.TabStopLeaderCharValues.Heavy)
        {
            return TabLeader.Heavy;
        }

        if (value == W.TabStopLeaderCharValues.MiddleDot)
        {
            return TabLeader.MiddleDot;
        }

        return TabLeader.None;
    }
}
