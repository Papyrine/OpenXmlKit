namespace OpenXmlKit.Word;

/// <summary>
/// Finds the package part an element was read from.
/// </summary>
/// <remarks>
/// A relationship id only means something against the part that declares it, and the part a piece
/// of content lives in is not always the main document — a hyperlink or a picture in a header
/// resolves against the header part. Walking to the root element and asking it, rather than
/// threading a part down through every view, is what keeps the views a single field wide.
/// </remarks>
static class PartLookup
{
    public static OpenXmlPart? Of(OpenXmlElement element)
    {
        var current = element;
        while (current.Parent is { } parent)
        {
            current = parent;
        }

        return (current as OpenXmlPartRootElement)?.OpenXmlPart;
    }
}
