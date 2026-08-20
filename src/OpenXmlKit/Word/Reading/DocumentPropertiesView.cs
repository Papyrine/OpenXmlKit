namespace OpenXmlKit.Word;

/// <summary>
/// A document's metadata, as read.
/// </summary>
public readonly struct DocumentPropertiesView
{
    readonly WordprocessingDocument package;

    internal DocumentPropertiesView(WordprocessingDocument package) =>
        this.package = package;

    /// <summary>
    /// The document title.
    /// </summary>
    public string? Title => package.PackageProperties.Title;

    /// <summary>
    /// What the document is about.
    /// </summary>
    public string? Subject => package.PackageProperties.Subject;

    /// <summary>
    /// The author.
    /// </summary>
    public string? Creator => package.PackageProperties.Creator;

    /// <summary>
    /// A longer description, shown as Comments in Word.
    /// </summary>
    public string? Description => package.PackageProperties.Description;

    /// <summary>
    /// Tags, separated however the consumer expects.
    /// </summary>
    public string? Keywords => package.PackageProperties.Keywords;

    /// <summary>
    /// The category the document belongs to.
    /// </summary>
    public string? Category => package.PackageProperties.Category;

    /// <summary>
    /// Who saved it last.
    /// </summary>
    public string? LastModifiedBy => package.PackageProperties.LastModifiedBy;

    /// <summary>
    /// When it was created.
    /// </summary>
    public DateTime? Created => package.PackageProperties.Created;

    /// <summary>
    /// When it was last saved.
    /// </summary>
    public DateTime? Modified => package.PackageProperties.Modified;

    /// <summary>
    /// The custom properties, which a template can bind fields to.
    /// </summary>
    public IEnumerable<KeyValuePair<string, string>> Custom
    {
        get
        {
            var properties = package.CustomFilePropertiesPart?.Properties;
            if (properties == null)
            {
                yield break;
            }

            foreach (var property in properties.Elements<CustomDocumentProperty>())
            {
                if (property.Name?.Value is { } name)
                {
                    yield return new(name, DocumentProperties.ValueOf(property));
                }
            }
        }
    }

    /// <summary>
    /// The value of a custom property, or null when the document has none by that name.
    /// </summary>
    public string? GetCustom(string name)
    {
        foreach (var property in Custom)
        {
            if (property.Key == name)
            {
                return property.Value;
            }
        }

        return null;
    }
}
