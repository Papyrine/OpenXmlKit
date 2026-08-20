namespace OpenXmlKit.Word;

/// <summary>
/// The six edges of a border box, as read.
/// </summary>
public interface IBordersView
{
    /// <inheritdoc cref="Borders.Top"/>
    IBorderView Top { get; }

    /// <inheritdoc cref="Borders.Bottom"/>
    IBorderView Bottom { get; }

    /// <inheritdoc cref="Borders.Left"/>
    IBorderView Left { get; }

    /// <inheritdoc cref="Borders.Right"/>
    IBorderView Right { get; }

    /// <inheritdoc cref="Borders.InsideHorizontal"/>
    IBorderView InsideHorizontal { get; }

    /// <inheritdoc cref="Borders.InsideVertical"/>
    IBorderView InsideVertical { get; }

    /// <inheritdoc cref="Borders.IsEmpty"/>
    bool IsEmpty { get; }
}
