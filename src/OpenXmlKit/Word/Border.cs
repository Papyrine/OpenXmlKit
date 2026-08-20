namespace OpenXmlKit.Word;

/// <summary>
/// One edge of a border box.
/// </summary>
public class Border :
    IBorderView
{
    /// <summary>
    /// The line to draw. Null leaves the edge unstated, so it inherits; <see cref="BorderStyle.None"/>
    /// states that there is no line, which is what overrides an inherited one.
    /// </summary>
    public BorderStyle? Style { get; set; }

    /// <summary>
    /// Line thickness. Word stores this in eighths of a point, which is why a
    /// <see cref="Length"/> and not a number: 4 eighths reads as "half a point" here and as
    /// "four points" almost everywhere else the same attribute name appears.
    /// </summary>
    public Length? Width { get; set; }

    /// <summary>
    /// The colour of the line.
    /// </summary>
    public Color? Color { get; set; }

    /// <summary>
    /// Gap between the border and the content it surrounds.
    /// </summary>
    public Length? Space { get; set; }

    /// <summary>
    /// Draws the border with a shadow, which Word renders as a thicker line on two sides.
    /// </summary>
    public bool Shadow { get; set; }

    /// <summary>
    /// Whether the border states anything. An empty border writes no element.
    /// </summary>
    public bool IsEmpty =>
        Style == null &&
        Width == null &&
        Color == null &&
        Space == null &&
        !Shadow;

    /// <summary>
    /// An independent copy.
    /// </summary>
    public Border Clone() =>
        new()
        {
            Style = Style,
            Width = Width,
            Color = Color,
            Space = Space,
            Shadow = Shadow
        };

    /// <summary>
    /// Overwrites every property with the other border.
    /// </summary>
    public void CopyFrom(Border other)
    {
        Style = other.Style;
        Width = other.Width;
        Color = other.Color;
        Space = other.Space;
        Shadow = other.Shadow;
    }

    /// <summary>
    /// Sets every aspect at once, for the common case of a plain line.
    /// </summary>
    public void Set(BorderStyle style, Length? width = null, Color? color = null)
    {
        Style = style;
        Width = width;
        Color = color;
    }

    internal void ApplyTo(W.BorderType element)
    {
        if (Style is { } style)
        {
            element.Val = style.ToOpenXml();
        }

        if (Width is { } width)
        {
            element.Size = (uint) width.EighthPoints;
        }

        if (Color is { } color)
        {
            element.Color = color.Value;
            if (color.IsTheme)
            {
                element.ThemeColor = color.Theme.ToOpenXml();
            }
        }

        if (Space is { } space)
        {
            element.Space = (uint) space.TotalPoints;
        }

        if (Shadow)
        {
            element.Shadow = true;
        }
    }

    internal void ReadFrom(W.BorderType element)
    {
        if (element.Val is { HasValue: true } value)
        {
            Style = Map.ToBorderStyle(value.Value);
        }

        if (element.Size is { HasValue: true } size)
        {
            Width = Length.FromEighthPoints(size.Value);
        }

        if (element.Color is { HasValue: true } color &&
            Word.Color.TryParse(color.Value, out var parsed))
        {
            Color = parsed;
        }

        if (element.Space is { HasValue: true } space)
        {
            Space = Length.FromPoints(space.Value);
        }

        Shadow = element.Shadow?.Value == true;
    }
}
