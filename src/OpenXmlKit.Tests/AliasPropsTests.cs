[TestFixture]
public class AliasPropsTests
{
    [Test]
    public void ShippedPropsFileIsUpToDate()
    {
        var path = PropsPath();
        var expected = AliasProps.Generate();
        var actual = File.Exists(path) ? File.ReadAllText(path) : "";

        if (Normalise(actual) == Normalise(expected))
        {
            return;
        }

        // Written rather than merely reported, so bringing it up to date is a matter of re-running
        // the suite and committing the diff. The list going stale is the one obvious way this
        // design rots, and it rots silently: nothing here breaks, a consumer's build does.
        File.WriteAllText(path, expected);
        Assert.Fail(
            $"The alias props file was out of date and has been regenerated at {path}. " +
            "Review the diff and re-run.");
    }

    [Test]
    public void EveryPublicTypeIsAliasedUnderBothPrefixes()
    {
        var props = AliasProps.Generate();

        foreach (var type in AliasProps.PublicTypes())
        {
            Assert.That(props, Does.Contain($"Alias=\"W{type.Name}\""), type.Name);
            Assert.That(props, Does.Contain($"Alias=\"Word{type.Name}\""), type.Name);
        }
    }

    [Test]
    public void TheTypesThatCollideWithTheSdkAreCovered()
    {
        // The names that make the aliases necessary in the first place. Losing one of these from
        // the public surface would be a breaking change worth noticing here rather than in a
        // consumer.
        var collisions = new[]
        {
            "Document", "Paragraph", "Run", "Table", "TableLook", "Border", "Borders", "Shading",
            "Style", "Styles", "Numbering", "Body", "Color", "Font", "TabStop", "TabStops"
        };

        var names = AliasProps.PublicTypes().Select(_ => _.Name).ToHashSet(StringComparer.Ordinal);
        Assert.That(names, Is.SupersetOf(collisions));
    }

    static string Normalise(string value) =>
        value.Replace("\r\n", "\n").TrimEnd();

    static string PropsPath()
    {
        var directory = new DirectoryInfo(TestContext.CurrentContext.TestDirectory);
        while (directory != null &&
               !Directory.Exists(Path.Combine(directory.FullName, "OpenXmlKit", "buildTransitive")))
        {
            directory = directory.Parent;
        }

        if (directory == null)
        {
            throw new DirectoryNotFoundException("Could not find the source directory from the test output directory.");
        }

        return Path.Combine(directory.FullName, "OpenXmlKit", "buildTransitive", "OpenXmlKit.props");
    }
}
