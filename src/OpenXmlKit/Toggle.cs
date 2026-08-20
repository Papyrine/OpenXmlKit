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

    /// <summary>
    /// Explicitly on.
    /// </summary>
    public static Toggle On => new(true);

    /// <summary>
    /// Explicitly off, overriding an inherited on.
    /// </summary>
    public static Toggle Off => new(false);

    /// <summary>
    /// Whether this states anything at all. <c>false</c> for <see cref="Inherit"/>.
    /// </summary>
    public bool IsSet => value.HasValue;

    /// <summary>
    /// Whether the value is explicitly on. False for both off and inherited.
    /// </summary>
    public bool IsOn => value == true;

    /// <summary>
    /// Whether the value is explicitly off, which is the state that cancels a style.
    /// </summary>
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

    /// <summary>
    /// Whether the two toggles are in the same one of the three states.
    /// </summary>
    public bool Equals(Toggle other) =>
        value == other.value;

    /// <summary>
    /// Whether the other object is a value of this type and equal to this one.
    /// </summary>
    public override bool Equals(object? obj) =>
        obj is Toggle other && Equals(other);

    /// <summary>
    /// A hash consistent with equality.
    /// </summary>
    public override int GetHashCode() =>
        value.GetHashCode();

    /// <summary>
    /// Whether the two toggles are in the same state.
    /// </summary>
    public static bool operator ==(Toggle left, Toggle right) => left.Equals(right);

    /// <summary>
    /// Whether the two toggles differ.
    /// </summary>
    public static bool operator !=(Toggle left, Toggle right) => !left.Equals(right);

    /// <summary>
    /// A readable form, for logs and debugging rather than for the file.
    /// </summary>
    public override string ToString() =>
        value switch
        {
            true => "On",
            false => "Off",
            null => "Inherit"
        };
}
