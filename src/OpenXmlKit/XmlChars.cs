#if NET8_0_OR_GREATER
using System.Buffers;
#endif

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

        var stripped = Strip(value.AsSpan());

        // Same memory means nothing was removed, so the original instance comes back rather than a
        // copy of itself.
        return stripped == value.AsSpan() ? value : stripped.ToString();
    }

    /// <summary>
    /// The given text with anything XML 1.0 forbids removed.
    /// </summary>
    /// <remarks>
    /// The span overload exists for callers holding a slice of a larger buffer — a parser's view
    /// over its input, typically — where going through a string would allocate one per call for
    /// text that virtually never needs changing. Clean input comes back as the same memory.
    /// </remarks>
    public static ReadOnlySpan<char> Strip(ReadOnlySpan<char> value)
    {
        if (value.Length == 0)
        {
            return value;
        }

        var suspect = IndexOfSuspect(value);
        if (suspect < 0)
        {
            return value;
        }

        // Everything before the first suspect is known good, so the per-character walk starts
        // there rather than at the beginning.
        return StripFrom(value, suspect);
    }

    /// <summary>
    /// Whether the text is already free of characters XML 1.0 forbids.
    /// </summary>
    public static bool IsValid(string value) =>
        ReferenceEquals(Strip(value), value);

    static ReadOnlySpan<char> StripFrom(ReadOnlySpan<char> value, int start)
    {
        // The builder stays null when nothing turns out to be invalid, which happens whenever the
        // scan above stopped on a surrogate that proves to be half of a valid pair — pair
        // detection needs the walk, so the scan cannot rule it out.
        StringBuilder? builder = null;
        var index = start;
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
                valid = !IsSuspect(character);
            }

            if (valid)
            {
                builder?.Append(value.Slice(index, advance));
            }
            else if (builder == null)
            {
                builder = new(value.Length);
                builder.Append(value[..index]);
            }

            index += advance;
        }

        if (builder == null)
        {
            return value;
        }

        return builder.ToString().AsSpan();
    }

    // Tab, line feed and carriage return are the three controls XML keeps. Everything below 0x20
    // besides those is forbidden, as are the two non-characters at the end of the BMP. Surrogates
    // fall outside both allowed ranges and so count as suspect, which is what sends a pair to the
    // walk that can recognise it.
    static bool IsSuspect(char character) =>
        character is not ('\t' or '\n' or '\r') &&
        character is not (>= (char) 0x20 and <= (char) 0xD7FF) &&
        character is not (>= (char) 0xE000 and <= (char) 0xFFFD);

#if NET8_0_OR_GREATER
    // A vectorised scan, so text with nothing to remove — which is nearly all of it — is rejected
    // in one pass rather than a character at a time.
    static readonly SearchValues<char> suspects = SearchValues.Create(BuildSuspects());

    static int IndexOfSuspect(ReadOnlySpan<char> value) =>
        value.IndexOfAny(suspects);

    static char[] BuildSuspects()
    {
        var result = new List<char>(2080);
        for (var character = 0; character <= 0xFFFF; character++)
        {
            if (IsSuspect((char) character))
            {
                result.Add((char) character);
            }
        }

        return [.. result];
    }
#else
    static int IndexOfSuspect(ReadOnlySpan<char> value)
    {
        for (var index = 0; index < value.Length; index++)
        {
            if (IsSuspect(value[index]))
            {
                return index;
            }
        }

        return -1;
    }
#endif
}
