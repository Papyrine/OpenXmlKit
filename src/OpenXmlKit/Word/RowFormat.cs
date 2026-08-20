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

    /// <summary>
    /// Grid columns to leave empty before the row's first cell.
    /// </summary>
    /// <remarks>
    /// How a ragged table is expressed: every row shares the table's one grid, so a row that
    /// starts part-way across states the gap rather than carrying placeholder cells.
    /// <see cref="WidthBefore"/> is what that gap measures, and Word needs both — the count to
    /// know which grid columns are skipped, the width to know how wide the skip is.
    /// </remarks>
    public int? GridBefore { get; set; }

    /// <summary>
    /// Grid columns to leave empty after the row's last cell.
    /// </summary>
    public int? GridAfter { get; set; }

    /// <summary>
    /// The width of the gap <see cref="GridBefore"/> leaves.
    /// </summary>
    public Width WidthBefore { get; set; } = Width.Auto;

    /// <summary>
    /// The width of the gap <see cref="GridAfter"/> leaves.
    /// </summary>
    public Width WidthAfter { get; set; } = Width.Auto;

    public bool IsEmpty =>
        Height == null &&
        HeightRule == RowHeightRule.Auto &&
        !IsHeader.IsSet &&
        !CantSplit.IsSet &&
        GridBefore == null &&
        GridAfter == null &&
        WidthBefore.IsAuto &&
        WidthAfter.IsAuto;

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
        GridBefore = other.GridBefore;
        GridAfter = other.GridAfter;
        WidthBefore = other.WidthBefore;
        WidthAfter = other.WidthAfter;
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
        // The order below is nonetheless the one Word itself writes, which keeps a diff against a
        // Word-saved file readable.
        if (GridBefore is { } gridBefore)
        {
            properties.Append(
                new W.GridBefore
                {
                    Val = gridBefore
                });
        }

        if (GridAfter is { } gridAfter)
        {
            properties.Append(
                new W.GridAfter
                {
                    Val = gridAfter
                });
        }

        if (!WidthBefore.IsAuto)
        {
            properties.Append(WidthElement.Of<W.WidthBeforeTableRow>(WidthBefore));
        }

        if (!WidthAfter.IsAuto)
        {
            properties.Append(WidthElement.Of<W.WidthAfterTableRow>(WidthAfter));
        }

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
        GridBefore = properties.GetFirstChild<W.GridBefore>()?.Val?.Value;
        GridAfter = properties.GetFirstChild<W.GridAfter>()?.Val?.Value;
        WidthBefore = WidthElement.Read(properties.GetFirstChild<W.WidthBeforeTableRow>());
        WidthAfter = WidthElement.Read(properties.GetFirstChild<W.WidthAfterTableRow>());

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
