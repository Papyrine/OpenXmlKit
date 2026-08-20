namespace OpenXmlKit.Word;

/// <summary>
/// A document's metadata, as read.
/// </summary>
public readonly struct DocumentPropertiesView
{
    readonly WordprocessingDocument package;

    internal DocumentPropertiesView(WordprocessingDocument package) =>
        this.package = package;

    public string? Title => package.PackageProperties.Title;
    public string? Subject => package.PackageProperties.Subject;
    public string? Creator => package.PackageProperties.Creator;
    public string? Description => package.PackageProperties.Description;
    public string? Keywords => package.PackageProperties.Keywords;
    public string? Category => package.PackageProperties.Category;
    public string? LastModifiedBy => package.PackageProperties.LastModifiedBy;
    public DateTime? Created => package.PackageProperties.Created;
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
