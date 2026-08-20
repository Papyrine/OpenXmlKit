namespace OpenXmlKit.Word;

/// <summary>
/// Helpers for bookmark names, which Word constrains more than a caller expects.
/// </summary>
public static class Bookmarks
{
    /// <summary>
    /// The longest name Word accepts.
    /// </summary>
    public const int MaxLength = 40;

    /// <summary>
    /// Turns arbitrary text into something Word will accept as a bookmark name.
    /// </summary>
    /// <remarks>
    /// A name must start with a letter, may contain only letters, digits and underscores, and is
    /// capped at forty characters. Word does not report a name that breaks those rules — it drops
    /// the bookmark, and every cross-reference to it renders as an error at the point a reader
    /// opens the document.
    /// <para>
    /// Sanitising is a last resort rather than a strategy: two titles can differ only in
    /// punctuation, or only past the fortieth character, and both collapse to the same name here.
    /// Where the names have to be unique, derive them from something positional instead.
    /// </para>
    /// </remarks>
    public static string Sanitise(string name)
    {
        var builder = new StringBuilder(name.Length);
        foreach (var character in name)
        {
            if (char.IsLetterOrDigit(character))
            {
                builder.Append(character);
                continue;
            }

            if (builder.Length > 0 &&
                builder[^1] != '_')
            {
                builder.Append('_');
            }
        }

        while (builder.Length > 0 &&
               !char.IsLetter(builder[0]))
        {
            builder.Remove(0, 1);
        }

        if (builder.Length == 0)
        {
            return "bookmark";
        }

        if (builder.Length > MaxLength)
        {
            builder.Length = MaxLength;
        }

        while (builder.Length > 0 &&
               builder[^1] == '_')
        {
            builder.Length--;
        }

        return builder.Length == 0 ? "bookmark" : builder.ToString();
    }

    /// <summary>
    /// Whether a name is one Word will keep.
    /// </summary>
    public static bool IsValid(string name)
    {
        if (name.Length is 0 or > MaxLength ||
            !char.IsLetter(name[0]))
        {
            return false;
        }

        foreach (var character in name)
        {
            if (!char.IsLetterOrDigit(character) &&
                character != '_')
            {
                return false;
            }
        }

        return true;
    }
}
