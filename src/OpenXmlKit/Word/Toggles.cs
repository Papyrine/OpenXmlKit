namespace OpenXmlKit.Word;

/// <summary>
/// Turns a <see cref="Toggle"/> into the element that states it, and reads one back.
/// </summary>
/// <remarks>
/// Two shapes exist in the schema for what is conceptually the same thing. CT_OnOff takes an
/// optional boolean and defaults to on when absent; CT_OnOffOnly takes an explicit on/off token.
/// Both are handled here so that callers only ever deal with <see cref="Toggle"/>.
/// <para>
/// A null return removes the child, which is what an inherited toggle wants. An off toggle
/// returns an element that says so, which is the only way to cancel formatting a style turned on.
/// </para>
/// </remarks>
static class Toggles
{
    public static T? OnOff<T>(Toggle toggle)
        where T : W.OnOffType, new()
    {
        if (!toggle.IsSet)
        {
            return null;
        }

        var element = new T();
        if (toggle.IsOff)
        {
            element.Val = false;
        }

        return element;
    }

    public static T? OnOffOnly<T>(Toggle toggle)
        where T : W.OnOffOnlyType, new()
    {
        if (!toggle.IsSet)
        {
            return null;
        }

        return new()
        {
            Val = toggle.IsOn ? W.OnOffOnlyValues.On : W.OnOffOnlyValues.Off
        };
    }

    public static Toggle Read(W.OnOffType? element)
    {
        if (element == null)
        {
            return Toggle.Inherit;
        }

        return element.Val?.Value != false;
    }

    public static Toggle Read(W.OnOffOnlyType? element)
    {
        if (element == null)
        {
            return Toggle.Inherit;
        }

        if (element.Val is not { HasValue: true } value)
        {
            return Toggle.On;
        }

        return value.Value == W.OnOffOnlyValues.On;
    }
}
