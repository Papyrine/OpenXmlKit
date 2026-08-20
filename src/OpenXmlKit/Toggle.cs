namespace OpenXmlKit;

/// <summary>
/// A Word toggle property — bold, italic, small-caps and the rest — which has three states, not two:
/// on, off, and say-nothing.
/// </summary>
/// <remarks>
/// <c>&lt;w:b/&gt;</c> means on and <c>&lt;w:b w:val="0"/&gt;</c> means off; emitting neither means
/// "whatever the style says". Modelling these as <c>bool</c> collapses off and say-nothing into one
/// value, and the result is that formatting can be turned on but never off: a run inside a bold
/// paragraph style has no way to be un-bolded, because the only thing a <c>false</c> can do is
/// decline to write the element.
/// <para>
/// <see cref="Inherit"/> is the default, so a format object left alone writes nothing and the style
/// hierarchy is left intact.
/// </para>
/// </remarks>
public readonly struct Toggle :
    IEquatable<Toggle>
{
    readonly bool? value;

    Toggle(bool? value) =>
        this.value = value;

    /// <summary>
    /// Say nothing, and take whatever the style hierarchy resolves to. The default.
    /// </summary>
    public static Toggle Inherit => default;

    public static Toggle On => new(true);

    /// <summary>
    /// Explicitly off, overriding an inherited on.
    /// </summary>
    public static Toggle Off => new(false);

    /// <summary>
    /// Whether this states anything at all. <c>false</c> for <see cref="Inherit"/>.
    /// </summary>
    public bool IsSet => value.HasValue;

    public bool IsOn => value == true;
    public bool IsOff => value == false;

    /// <summary>
    /// The underlying state, or <c>null</c> when inheriting.
    /// </summary>
    public bool? Value => value;

    /// <summary>
    /// Reads a <c>bool</c> as on/off and a <c>null</c> as inherit, so
    /// <c>Bold = true</c> and <c>Bold = null</c> both read naturally.
    /// </summary>
    public static implicit operator Toggle(bool? value) =>
        new(value);

    /// <summary>
    /// Reads as "is this on", so <see cref="Inherit"/> and <see cref="Off"/> are both false.
    /// Use <see cref="IsSet"/> to tell them apart.
    /// </summary>
    public static implicit operator bool(Toggle toggle) =>
        toggle.value == true;

    public bool Equals(Toggle other) =>
        value == other.value;

    public override bool Equals(object? obj) =>
        obj is Toggle other && Equals(other);

    public override int GetHashCode() =>
        value.GetHashCode();

    public static bool operator ==(Toggle left, Toggle right) => left.Equals(right);
    public static bool operator !=(Toggle left, Toggle right) => !left.Equals(right);

    public override string ToString() =>
        value switch
        {
            true => "On",
            false => "Off",
            null => "Inherit"
        };
}
