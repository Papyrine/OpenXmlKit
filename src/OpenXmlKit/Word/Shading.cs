namespace OpenXmlKit.Word;

/// <summary>
/// A background fill, on a run, a paragraph, or a table cell.
/// </summary>
/// <remarks>
/// The underlying element is a pattern fill — a foreground colour, a background colour, and a
/// texture blending them — which means the plain solid colour everybody actually wants is spelled
/// as a <c>clear</c> pattern over a background fill, and getting that combination wrong produces
/// either nothing or black. <see cref="BackgroundColor"/> is the whole of it for a solid fill;
/// the pattern properties are there for the rare case that needs them.
/// <para>
/// A fill is worth having as its own concept because a Word style cannot carry one per cell — it
/// is the reason the stocktake report writes shading inline while taking everything else from a
/// named table style.
/// </para>
/// </remarks>
public class Shading :
    IShadingView
{
    /// <summary>
    /// The solid colour to fill with.
    /// </summary>
    public Color? BackgroundColor { get; set; }

    /// <summary>
    /// The pattern colour, for the patterned fills. Unused by a solid fill.
    /// </summary>
    public Color? PatternColor { get; set; }

    /// <summary>
    /// The blend between the two colours. Defaults to <c>clear</c>, which shows
    /// <see cref="BackgroundColor"/> alone.
    /// </summary>
    public ShadingPattern Pattern { get; set; } = ShadingPattern.Clear;

    public bool IsEmpty =>
        BackgroundColor == null &&
        PatternColor == null &&
        Pattern == ShadingPattern.Clear;

    public Shading Clone() =>
        new()
        {
            BackgroundColor = BackgroundColor,
            PatternColor = PatternColor,
            Pattern = Pattern
        };

    public void CopyFrom(Shading other)
    {
        BackgroundColor = other.BackgroundColor;
        PatternColor = other.PatternColor;
        Pattern = other.Pattern;
    }

    internal W.Shading? ToOpenXml()
    {
        if (IsEmpty)
        {
            return null;
        }

        var shading = new W.Shading
        {
            Val = Pattern == ShadingPattern.Solid
                ? W.ShadingPatternValues.Solid
                : W.ShadingPatternValues.Clear
        };

        if (BackgroundColor is { } background)
        {
            shading.Fill = background.Value;
            if (background.IsTheme)
            {
                shading.ThemeFill = background.Theme.ToOpenXml();
            }
        }
        else
        {
            // Word wants a fill attribute present; "auto" is the no-opinion value.
            shading.Fill = "auto";
        }

        shading.Color = PatternColor?.Value ?? "auto";
        return shading;
    }

    internal void ReadFrom(W.Shading? shading)
    {
        if (shading == null)
        {
            return;
        }

        if (shading.Fill is { HasValue: true } fill &&
            Color.TryParse(fill.Value, out var background))
        {
            if (!background.IsAuto)
            {
                BackgroundColor = background;
            }
        }

        if (shading.Color is { HasValue: true } color &&
            Color.TryParse(color.Value, out var pattern))
        {
            if (!pattern.IsAuto)
            {
                PatternColor = pattern;
            }
        }

        if (shading.Val is { HasValue: true } value &&
            value.Value == W.ShadingPatternValues.Solid)
        {
            Pattern = ShadingPattern.Solid;
        }
    }
}
