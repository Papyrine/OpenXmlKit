namespace OpenXmlKit.Word;

/// <summary>
/// Turns a <see cref="Width"/> into the value-plus-type pair the schema wants, and back.
/// </summary>
/// <remarks>
/// Three elements carry a width in the same shape — the table, the cell, and a cell margin — and
/// each is a distinct SDK type over the same CT_TblWidth base. Converting once here is what keeps
/// the fiftieths-of-a-percent encoding from being restated at each of them.
/// </remarks>
static class WidthElement
{
    public static W.TableWidth TableWidth(Width width) =>
        Of<W.TableWidth>(width);

    public static W.TableCellWidth CellWidth(Width width) =>
        Of<W.TableCellWidth>(width);

    /// <summary>
    /// The same conversion for any of the CT_TblWidth elements that has no named helper here.
    /// </summary>
    public static T Of<T>(Width width)
        where T : W.TableWidthType, new()
    {
        var element = new T();
        Fill(element, width);
        return element;
    }

    static void Fill(W.TableWidthType element, Width width)
    {
        switch (width.Unit)
        {
            case WidthUnit.Percent:
                element.Width = width.FiftiethsOfAPercent.ToString(CultureInfo.InvariantCulture);
                element.Type = W.TableWidthUnitValues.Pct;
                return;
            case WidthUnit.Absolute:
                element.Width = width.AsLength.Twips.ToString(CultureInfo.InvariantCulture);
                element.Type = W.TableWidthUnitValues.Dxa;
                return;
            default:
                element.Width = "0";
                element.Type = W.TableWidthUnitValues.Auto;
                return;
        }
    }

    public static Width Read(W.TableWidthType? element)
    {
        if (element?.Width is not { HasValue: true } raw ||
            !double.TryParse(raw.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value))
        {
            return Width.Auto;
        }

        var type = element.Type?.Value;
        if (type == null)
        {
            return Width.Auto;
        }

        if (type.Value == W.TableWidthUnitValues.Pct)
        {
            return Width.Percent(value / 50);
        }

        if (type.Value == W.TableWidthUnitValues.Dxa)
        {
            return Width.FromTwips(value);
        }

        return Width.Auto;
    }
}
