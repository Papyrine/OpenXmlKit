namespace OpenXmlKit.Word;

/// <summary>
/// A document's metadata: the built-in core properties, plus whatever custom properties it carries.
/// </summary>
public class DocumentProperties
{
    readonly WordprocessingDocument package;

    internal DocumentProperties(WordprocessingDocument package) =>
        this.package = package;

    /// <summary>
    /// The document title.
    /// </summary>
    public string? Title
    {
        get => package.PackageProperties.Title;
        set => package.PackageProperties.Title = value;
    }

    /// <summary>
    /// What the document is about.
    /// </summary>
    public string? Subject
    {
        get => package.PackageProperties.Subject;
        set => package.PackageProperties.Subject = value;
    }

    /// <summary>
    /// The author.
    /// </summary>
    public string? Creator
    {
        get => package.PackageProperties.Creator;
        set => package.PackageProperties.Creator = value;
    }

    /// <summary>
    /// A longer description, shown as Comments in Word.
    /// </summary>
    public string? Description
    {
        get => package.PackageProperties.Description;
        set => package.PackageProperties.Description = value;
    }

    /// <summary>
    /// Tags, separated however the consumer expects.
    /// </summary>
    public string? Keywords
    {
        get => package.PackageProperties.Keywords;
        set => package.PackageProperties.Keywords = value;
    }

    /// <summary>
    /// The category the document belongs to.
    /// </summary>
    public string? Category
    {
        get => package.PackageProperties.Category;
        set => package.PackageProperties.Category = value;
    }

    /// <summary>
    /// Who saved it last.
    /// </summary>
    public string? LastModifiedBy
    {
        get => package.PackageProperties.LastModifiedBy;
        set => package.PackageProperties.LastModifiedBy = value;
    }

    /// <summary>
    /// When the document was created.
    /// </summary>
    /// <remarks>
    /// Worth setting explicitly on a generated document: left alone it is the moment of generation,
    /// which makes two runs over the same data produce different bytes.
    /// </remarks>
    public DateTime? Created
    {
        get => package.PackageProperties.Created;
        set => package.PackageProperties.Created = value;
    }

    /// <summary>
    /// When it was last saved.
    /// </summary>
    public DateTime? Modified
    {
        get => package.PackageProperties.Modified;
        set => package.PackageProperties.Modified = value;
    }

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
                    yield return new(name, ValueOf(property));
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

    /// <summary>
    /// Sets a custom property, replacing any of the same name.
    /// </summary>
    public DocumentProperties SetCustom(string name, string value)
    {
        var part = package.CustomFilePropertiesPart ?? package.AddCustomFilePropertiesPart();
        part.Properties ??= new();
        var properties = part.Properties;

        foreach (var existing in properties.Elements<CustomDocumentProperty>().ToList())
        {
            if (existing.Name?.Value == name)
            {
                existing.Remove();
            }
        }

        properties.AppendChild(
            new CustomDocumentProperty(
                new DocumentFormat.OpenXml.VariantTypes.VTLPWSTR(value))
            {
                Name = name,
                // The format id is fixed for user-defined properties, and the ids run from 2 —
                // 0 and 1 are reserved.
                FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                PropertyId = properties.Elements<CustomDocumentProperty>().Count() + 2
            });
        return this;
    }

    internal static string ValueOf(CustomDocumentProperty property) =>
        property.VTLPWSTR?.Text ??
        property.VTBString?.Text ??
        property.VTInt32?.Text ??
        property.VTDouble?.Text ??
        property.VTBool?.Text ??
        property.VTFileTime?.Text ??
        property.InnerText;
}
