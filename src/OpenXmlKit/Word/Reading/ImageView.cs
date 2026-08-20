using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OpenXmlKit.Word;

/// <summary>
/// A picture in a document being read.
/// </summary>
/// <remarks>
/// The drawing tree a picture lives in restates its size at three levels and holds the image data
/// in a separate part. This reads the outermost of the three, which is the one that governs how
/// large the picture is drawn, and resolves the part on request rather than on construction — a
/// caller counting pictures should not pay for decoding them.
/// </remarks>
public readonly struct ImageView
{
    readonly W.Drawing element;

    internal ImageView(W.Drawing element) =>
        this.element = element;

    /// <summary>
    /// Whether the picture floats beside the text rather than sitting in the line.
    /// </summary>
    public bool IsFloating =>
        element.GetFirstChild<DW.Anchor>() != null;

    /// <summary>
    /// The width the picture is drawn at, which is not necessarily the image's own width.
    /// </summary>
    public Length Width =>
        Extent?.Cx is { } cx ? Length.FromEmu(cx) : Length.Zero;

    /// <summary>
    /// The height the picture is drawn at, which is not necessarily the image own height.
    /// </summary>
    public Length Height =>
        Extent?.Cy is { } cy ? Length.FromEmu(cy) : Length.Zero;

    /// <summary>
    /// The alternative text, which is what a screen reader announces.
    /// </summary>
    public string? Description => Properties?.Description?.Value;

    /// <summary>
    /// The name the drawing carries, which Word shows in the selection pane.
    /// </summary>
    public string? Name => Properties?.Name?.Value;

    /// <summary>
    /// The MIME type of the stored image, or null if the part cannot be reached.
    /// </summary>
    public string? ContentType => Part?.ContentType;

    /// <summary>
    /// The encoded image, exactly as it is stored in the package.
    /// </summary>
    /// <returns>
    /// Null when the drawing references no image part — a chart or a shape rather than a picture.
    /// </returns>
    public byte[]? GetBytes()
    {
        if (Part is not { } part)
        {
            return null;
        }

        using var source = part.GetStream();
        using var buffer = new MemoryStream();
        source.CopyTo(buffer);
        return buffer.ToArray();
    }

    /// <summary>
    /// The underlying OpenXML element, for anything this view does not expose.
    /// </summary>
    public W.Drawing ToOpenXml() =>
        element;

    ImagePart? Part
    {
        get
        {
            if (element.Descendants<A.Blip>().FirstOrDefault()?.Embed?.Value is not { } id ||
                PartLookup.Of(element) is not { } container)
            {
                return null;
            }

            return container.GetPartById(id) as ImagePart;
        }
    }

    DW.Extent? Extent =>
        element.GetFirstChild<DW.Inline>()?.Extent ??
        element.GetFirstChild<DW.Anchor>()?.Extent;

    // Inline exposes a typed DocProperties and Anchor does not, so both go through the untyped
    // lookup rather than one of each.
    DW.DocProperties? Properties =>
        element.GetFirstChild<DW.Inline>()?.GetFirstChild<DW.DocProperties>() ??
        element.GetFirstChild<DW.Anchor>()?.GetFirstChild<DW.DocProperties>();
}
