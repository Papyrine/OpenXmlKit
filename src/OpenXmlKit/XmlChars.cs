namespace OpenXmlKit;

/// <summary>
/// Removes the characters XML 1.0 forbids, which the SDK will not write and does not escape.
/// </summary>
/// <remarks>
/// Most C0 controls, unpaired surrogates and the two non-characters cannot appear in an XML
/// document at all. Nothing rejects them on the way in — <c>new Text("\u0001")</c> is a perfectly
/// ordinary object — so a value carrying one travels all the way to <c>Save</c> and throws there,
/// with nothing to say which of the document's strings was at fault. Text that came from a user, a
/// database or an HTML entity is exactly where they come from.
/// <para>
/// This is applied to every string the build API turns into a <c>w:t</c>, so a caller never has to
/// know. It is public because the same problem arises in any other part written from the same
/// data, and because a caller that assembles elements through the escape hatch has to do it
/// itself.
/// </para>
/// <para>
/// The SDK is expected to take this over eventually, by escaping such characters with the OOXML
/// <c>_xHHHH_</c> convention: https://github.com/dotnet/Open-XML-SDK/issues/1532.
/// </para>
/// </remarks>
public static class XmlChars
{
    /// <summary>
    /// The given text with anything XML 1.0 forbids removed.
    /// </summary>
    /// <remarks>
    /// Returns the original instance when there is nothing to remove, which is the overwhelmingly
    /// common case, so ordinary text costs one scan and no allocation.
    /// </remarks>
    public static string Strip(string value)
    {
        if (value.Length == 0)
        {
            return value;
        }

        StringBuilder? builder = null;
        var index = 0;
        while (index < value.Length)
        {
            var character = value[index];
            int advance;
            bool valid;

            // A surrogate is only legal as a matched pair, and the pair has to be stepped over
            // together — testing either half on its own would reject both.
            if (char.IsHighSurrogate(character) &&
                index + 1 < value.Length &&
                char.IsLowSurrogate(value[index + 1]))
            {
                advance = 2;
                valid = true;
            }
            else if (char.IsSurrogate(character))
            {
                advance = 1;
                valid = false;
            }
            else
            {
                advance = 1;
                valid = IsAllowed(character);
            }

            if (valid)
            {
                builder?.Append(value, index, advance);
            }
            else if (builder == null)
            {
                builder = new(value.Length);
                builder.Append(value, 0, index);
            }

            index += advance;
        }

        return builder?.ToString() ?? value;
    }

    /// <summary>
    /// Whether the text is already free of characters XML 1.0 forbids.
    /// </summary>
    public static bool IsValid(string value) =>
        ReferenceEquals(Strip(value), value);

    // Tab, line feed and carriage return are the three controls XML keeps. Everything below 0x20
    // besides those is forbidden, as are the two non-characters at the end of the BMP.
    static bool IsAllowed(char character) =>
        character is '\t' or '\n' or '\r' ||
        character is >= (char) 0x20 and <= (char) 0xD7FF ||
        character is >= (char) 0xE000 and <= (char) 0xFFFD;
}
