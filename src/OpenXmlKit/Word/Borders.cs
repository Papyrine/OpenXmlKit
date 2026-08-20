namespace OpenXmlKit.Word;

/// <summary>
/// The six edges a border box can have: the four sides, plus the two interior grids a table draws
/// between its cells.
/// </summary>
/// <remarks>
/// The whole-collection setters — <see cref="Style"/>, <see cref="Width"/>, <see cref="Color"/> —
/// fan out to every edge, so the overwhelmingly common "single line all round" is one statement
/// rather than six near-identical objects. That shape is lifted from Aspose, and it is what
/// replaces the six-fold repetition in the estate.
/// </remarks>
public class Borders
{
    public Border Top { get; } = new();
    public Border Bottom { get; } = new();
    public Border Left { get; } = new();
    public Border Right { get; } = new();

    /// <summary>
    /// The horizontal lines between rows. Tables only.
    /// </summary>
    public Border InsideHorizontal { get; } = new();

    /// <summary>
    /// The vertical lines between columns. Tables only.
    /// </summary>
    public Border InsideVertical { get; } = new();

    public bool IsEmpty =>
        Top.IsEmpty &&
        Bottom.IsEmpty &&
        Left.IsEmpty &&
        Right.IsEmpty &&
        InsideHorizontal.IsEmpty &&
        InsideVertical.IsEmpty;

    /// <summary>
    /// Sets the line style on every edge.
    /// </summary>
    public BorderStyle? Style
    {
        set
        {
            foreach (var border in All())
            {
                border.Style = value;
            }
        }
    }

    /// <summary>
    /// Sets the line thickness on every edge.
    /// </summary>
    public Length? Width
    {
        set
        {
            foreach (var border in All())
            {
                border.Width = value;
            }
        }
    }

    /// <summary>
    /// Sets the colour of every edge.
    /// </summary>
    public Color? Color
    {
        set
        {
            foreach (var border in All())
            {
                border.Color = value;
            }
        }
    }

    /// <summary>
    /// Draws the same line on all six edges.
    /// </summary>
    public void SetAll(BorderStyle style, Length? width = null, Color? color = null)
    {
        foreach (var border in All())
        {
            border.Set(style, width, color);
        }
    }

    /// <summary>
    /// Draws the same line on the four outside edges, leaving any interior grid alone.
    /// </summary>
    public void SetOutside(BorderStyle style, Length? width = null, Color? color = null)
    {
        Top.Set(style, width, color);
        Bottom.Set(style, width, color);
        Left.Set(style, width, color);
        Right.Set(style, width, color);
    }

    /// <summary>
    /// Draws the same line on the two interior grids, leaving the outside alone.
    /// </summary>
    public void SetInside(BorderStyle style, Length? width = null, Color? color = null)
    {
        InsideHorizontal.Set(style, width, color);
        InsideVertical.Set(style, width, color);
    }

    public IEnumerable<Border> All()
    {
        yield return Top;
        yield return Bottom;
        yield return Left;
        yield return Right;
        yield return InsideHorizontal;
        yield return InsideVertical;
    }

    public Borders Clone()
    {
        var clone = new Borders();
        clone.CopyFrom(this);
        return clone;
    }

    public void CopyFrom(Borders other)
    {
        Top.CopyFrom(other.Top);
        Bottom.CopyFrom(other.Bottom);
        Left.CopyFrom(other.Left);
        Right.CopyFrom(other.Right);
        InsideHorizontal.CopyFrom(other.InsideHorizontal);
        InsideVertical.CopyFrom(other.InsideVertical);
    }

    // Each container has its own set of border element types for the same six edges, so the three
    // builders below differ only in which types they instantiate. Assignment is through the typed
    // properties, which is what puts the children in CT_*Borders sequence without this code
    // knowing what that sequence is.
    internal W.TableBorders? ToTableBorders()
    {
        if (IsEmpty)
        {
            return null;
        }

        var borders = new W.TableBorders();
        Fill(Top, () => borders.TopBorder = new(), _ => borders.TopBorder = _);
        Fill(Left, () => borders.LeftBorder = new(), _ => borders.LeftBorder = _);
        Fill(Bottom, () => borders.BottomBorder = new(), _ => borders.BottomBorder = _);
        Fill(Right, () => borders.RightBorder = new(), _ => borders.RightBorder = _);
        Fill(InsideHorizontal, () => borders.InsideHorizontalBorder = new(), _ => borders.InsideHorizontalBorder = _);
        Fill(InsideVertical, () => borders.InsideVerticalBorder = new(), _ => borders.InsideVerticalBorder = _);
        return borders;
    }

    internal W.TableCellBorders? ToCellBorders()
    {
        if (IsEmpty)
        {
            return null;
        }

        var borders = new W.TableCellBorders();
        Fill(Top, () => borders.TopBorder = new(), _ => borders.TopBorder = _);
        Fill(Left, () => borders.LeftBorder = new(), _ => borders.LeftBorder = _);
        Fill(Bottom, () => borders.BottomBorder = new(), _ => borders.BottomBorder = _);
        Fill(Right, () => borders.RightBorder = new(), _ => borders.RightBorder = _);
        Fill(InsideHorizontal, () => borders.InsideHorizontalBorder = new(), _ => borders.InsideHorizontalBorder = _);
        Fill(InsideVertical, () => borders.InsideVerticalBorder = new(), _ => borders.InsideVerticalBorder = _);
        return borders;
    }

    internal W.ParagraphBorders? ToParagraphBorders()
    {
        if (Top.IsEmpty &&
            Bottom.IsEmpty &&
            Left.IsEmpty &&
            Right.IsEmpty)
        {
            return null;
        }

        var borders = new W.ParagraphBorders();
        Fill(Top, () => borders.TopBorder = new(), _ => borders.TopBorder = _);
        Fill(Left, () => borders.LeftBorder = new(), _ => borders.LeftBorder = _);
        Fill(Bottom, () => borders.BottomBorder = new(), _ => borders.BottomBorder = _);
        Fill(Right, () => borders.RightBorder = new(), _ => borders.RightBorder = _);
        return borders;
    }

    static void Fill<T>(Border border, Func<T> create, Action<T> assign)
        where T : W.BorderType, new()
    {
        if (border.IsEmpty)
        {
            return;
        }

        var element = new T();
        border.ApplyTo(element);
        assign(element);
    }

    internal void ReadFrom(OpenXmlElement? container)
    {
        if (container == null)
        {
            return;
        }

        foreach (var child in container.ChildElements)
        {
            if (child is not W.BorderType border)
            {
                continue;
            }

            var target = child switch
            {
                W.TopBorder => Top,
                W.BottomBorder => Bottom,
                W.LeftBorder or W.StartBorder => Left,
                W.RightBorder or W.EndBorder => Right,
                W.InsideHorizontalBorder => InsideHorizontal,
                W.InsideVerticalBorder => InsideVertical,
                _ => null
            };

            target?.ReadFrom(border);
        }
    }
}
