namespace OpenXmlKit.Word;

/// <summary>
/// The list definitions a document carries, and the ids they are allocated.
/// </summary>
/// <remarks>
/// Id allocation is the whole reason this exists as a type. Both ids are document-scoped integers
/// that have to be unique and have to be found before they can be extended — a document opened and
/// added to must continue past the ids already in it, or the new lists silently take over the
/// numbering of the old. Allocating here means no caller has to know that.
/// </remarks>
public class Numbering
{
    readonly Document document;
    readonly List<Pending> pending = [];
    int nextAbstractId = -1;
    int nextNumberingId = -1;

    internal Numbering(Document document) =>
        this.document = document;

    readonly record struct Pending(ListDefinition Definition, bool CreateAbstract);

    /// <summary>
    /// A bulleted list, with the nine levels Word gives a default bullet list: disc, circle,
    /// square, repeating, each indented half an inch further than the last.
    /// </summary>
    public ListDefinition AddBullet()
    {
        var levels = new List<ListLevel>();
        for (var depth = 0; depth < levelCount; depth++)
        {
            var level = new ListLevel(depth)
            {
                Format = NumberFormat.Bullet,
                Text = BulletGlyph(depth),
                Indent = Length.FromInches(0.5 * (depth + 1)),
                Hanging = Length.FromInches(0.25),
                Font =
                {
                    // The glyph comes from the font named here rather than from the paragraph's own, which
                    // is why a bullet survives a document whose body font has no such character.
                    Name = BulletFont(depth)
                }
            };
            levels.Add(level);
        }

        return Register(levels);
    }

    /// <summary>
    /// A numbered list, with the nine levels Word gives a default numbered list: the chosen format
    /// at the top, then lower letter and lower roman, repeating.
    /// </summary>
    public ListDefinition AddNumbered(NumberFormat format = NumberFormat.Decimal)
    {
        var levels = new List<ListLevel>();
        for (var depth = 0; depth < levelCount; depth++)
        {
            var levelFormat = depth == 0
                ? format
                : (depth % 3) switch
                {
                    1 => NumberFormat.LowerLetter,
                    2 => NumberFormat.LowerRoman,
                    _ => NumberFormat.Decimal
                };

            levels.Add(
                new(depth)
                {
                    Format = levelFormat,
                    // %n is the placeholder for the number at level n, counting from one.
                    Text = $"%{depth + 1}.",
                    Indent = Length.FromInches(0.5 * (depth + 1)),
                    Hanging = Length.FromInches(0.25),
                    Alignment = levelFormat == NumberFormat.LowerRoman
                        ? ListLevelAlignment.Right
                        : ListLevelAlignment.Left
                });
        }

        return Register(levels);
    }

    /// <summary>
    /// A list numbered as <c>1.1.1</c>, each level including the ones above it.
    /// </summary>
    public ListDefinition AddOutline()
    {
        var levels = new List<ListLevel>();
        for (var depth = 0; depth < levelCount; depth++)
        {
            var text = new StringBuilder();
            for (var part = 0; part <= depth; part++)
            {
                text.Append('%').Append((part + 1).ToString(CultureInfo.InvariantCulture));
                if (part < depth)
                {
                    text.Append('.');
                }
            }

            levels.Add(
                new(depth)
                {
                    Format = NumberFormat.Decimal,
                    Text = text.ToString(),
                    Indent = Length.FromInches(0.5 * (depth + 1)),
                    Hanging = Length.FromInches(0.5)
                });
        }

        return Register(levels);
    }

    /// <summary>
    /// A list built from scratch.
    /// </summary>
    public ListDefinition Add(Action<ListDefinition> configure)
    {
        var levels = new List<ListLevel>();
        for (var depth = 0; depth < levelCount; depth++)
        {
            levels.Add(new(depth));
        }

        var definition = Register(levels);
        configure(definition);
        return definition;
    }

    /// <summary>
    /// A second list that looks exactly like an existing one but counts from the start again.
    /// </summary>
    /// <remarks>
    /// This is what a document with two separate numbered lists needs. Reusing the same definition
    /// for both makes the second continue the first — the numbering belongs to the instance, not to
    /// the appearance.
    /// </remarks>
    public ListDefinition Restart(ListDefinition definition)
    {
        var copy = new ListDefinition(definition.AbstractId, AllocateNumberingId(), [.. definition.Levels]);
        pending.Add(new(copy, CreateAbstract: false));
        return copy;
    }

    ListDefinition Register(List<ListLevel> levels)
    {
        var definition = new ListDefinition(AllocateAbstractId(), AllocateNumberingId(), levels);
        pending.Add(new(definition, CreateAbstract: true));
        return definition;
    }

    const int levelCount = 9;

    static string BulletGlyph(int depth) =>
        (depth % 3) switch
        {
            // Symbol 0xF0B7 is a filled disc, Courier New "o" is a hollow circle, and Wingdings
            // 0xF0A7 is a filled square. These are the three Word cycles through.
            0 => "",
            1 => "o",
            _ => ""
        };

    static string BulletFont(int depth) =>
        (depth % 3) switch
        {
            0 => "Symbol",
            1 => "Courier New",
            _ => "Wingdings"
        };

    int AllocateAbstractId()
    {
        if (nextAbstractId < 0)
        {
            nextAbstractId = Existing()
                .Elements<W.AbstractNum>()
                .Select(_ => _.AbstractNumberId?.Value ?? 0)
                .DefaultIfEmpty(-1)
                .Max() + 1;
        }

        return nextAbstractId++;
    }

    int AllocateNumberingId()
    {
        if (nextNumberingId < 0)
        {
            nextNumberingId = Existing()
                .Elements<W.NumberingInstance>()
                .Select(_ => _.NumberID?.Value ?? 0)
                .DefaultIfEmpty(0)
                .Max() + 1;
        }

        return nextNumberingId++;
    }

    W.Numbering Existing() =>
        document.MainPart.NumberingDefinitionsPart?.Numbering ?? new();

    const string numberingRelationshipId = "rNumbering";

    internal void Save()
    {
        if (pending.Count == 0)
        {
            return;
        }

        var part = document.MainPart.NumberingDefinitionsPart ??
                   document.MainPart.AddNewPart<NumberingDefinitionsPart>(numberingRelationshipId);
        var root = part.Numbering ??= new();

        foreach (var entry in pending.Where(_ => _.CreateAbstract))
        {
            root.AppendChild(BuildAbstract(entry.Definition));
        }

        // Every abstractNum has to precede every num in the part: the schema sequences them, and
        // Word rejects the file otherwise. Since abstracts are written first and instances second,
        // and both append, the order falls out — but only because the two loops are separate.
        foreach (var entry in pending)
        {
            root.AppendChild(
                new W.NumberingInstance(
                    new W.AbstractNumId
                    {
                        Val = entry.Definition.AbstractId
                    })
                {
                    NumberID = entry.Definition.NumberingId
                });
        }

        pending.Clear();
        root.Save();
    }

    static W.AbstractNum BuildAbstract(ListDefinition definition)
    {
        var element = new W.AbstractNum
        {
            AbstractNumberId = definition.AbstractId,
            MultiLevelType = new()
            {
                Val = W.MultiLevelValues.HybridMultilevel
            }
        };

        foreach (var level in definition.Levels)
        {
            element.AppendChild(level.ToOpenXml());
        }

        return element;
    }
}
