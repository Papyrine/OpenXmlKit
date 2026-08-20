namespace OpenXmlKit.Word;

/// <summary>
/// Formatting for one table row.
/// </summary>
public class RowFormat :
    IRowFormatView
{
    public Length? Height { get; set; }

    public RowHeightRule HeightRule { get; set; } = RowHeightRule.Auto;

    /// <summary>
    /// Repeats this row at the top of every page the table spans.
    /// </summary>
    /// <remarks>
    /// Only meaningful on the rows at the top of the table, and worth setting on the header of any
    /// table long enough to break: without it, every page after the first is unlabelled columns of
    /// data.
    /// </remarks>
    public Toggle IsHeader { get; set; }

    /// <summary>
    /// Keeps the row from being split across a page boundary.
    /// </summary>
    public Toggle CantSplit { get; set; }

    public bool IsEmpty =>
        Height == null &&
        HeightRule == RowHeightRule.Auto &&
        !IsHeader.IsSet &&
        !CantSplit.IsSet;

    public RowFormat Clone()
    {
        var clone = new RowFormat();
        clone.CopyFrom(this);
        return clone;
    }

    public void CopyFrom(RowFormat other)
    {
        Height = other.Height;
        HeightRule = other.HeightRule;
        IsHeader = other.IsHeader;
        CantSplit = other.CantSplit;
    }

    public void Clear() =>
        CopyFrom(new());

    internal W.TableRowProperties? ToProperties()
    {
        if (IsEmpty)
        {
            return null;
        }

        var properties = new W.TableRowProperties();

        // CT_TrPr is a repeating choice rather than a sequence, so its children have no fixed
        // order to get wrong and no typed properties to assign through. Append is correct here.
        if (CantSplit.IsSet)
        {
            properties.Append(Toggles.OnOffOnly<W.CantSplit>(CantSplit)!);
        }

        if (Height is { } height)
        {
            properties.Append(
                new W.TableRowHeight
                {
                    Val = (uint) height.Twips,
                    HeightType = HeightRule.ToOpenXml()
                });
        }
        else if (HeightRule != RowHeightRule.Auto)
        {
            properties.Append(
                new W.TableRowHeight
                {
                    HeightType = HeightRule.ToOpenXml()
                });
        }

        if (IsHeader.IsSet)
        {
            properties.Append(Toggles.OnOffOnly<W.TableHeader>(IsHeader)!);
        }

        return properties;
    }

    internal void ReadFrom(W.TableRowProperties? properties)
    {
        if (properties == null)
        {
            return;
        }

        CantSplit = Toggles.Read(properties.GetFirstChild<W.CantSplit>());
        IsHeader = Toggles.Read(properties.GetFirstChild<W.TableHeader>());

        if (properties.GetFirstChild<W.TableRowHeight>() is not { } height)
        {
            return;
        }

        if (height.Val is { HasValue: true } value)
        {
            Height = Length.FromTwips(value.Value);
        }

        if (height.HeightType is { HasValue: true } rule)
        {
            HeightRule = Map.ToRowHeightRule(rule.Value);
        }
    }
}
