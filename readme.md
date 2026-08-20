# <img src="/src/icon.png" height="30px"> OpenXmlKit


[![Build status](https://img.shields.io/appveyor/build/SimonCropp/OpenXmlKit)](https://ci.appveyor.com/project/SimonCropp/OpenXmlKit)
[![NuGet Status](https://img.shields.io/nuget/v/OpenXmlKit.svg)](https://www.nuget.org/packages/OpenXmlKit/)

OpenXmlKit is an ergonomic wrapper over [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK)
for building and reading Word documents. It wraps the SDK rather than replacing it, so anything it
does not model is still reachable and a partial migration onto it is always possible.


## Why

The OpenXML SDK is a faithful projection of the file format, which makes it precise and makes it
hostile. Four things in particular cost time on every document:

**Schema child order is the caller's problem.** `w:rPr` has to list its children in the order the
schema declares — `rFonts`, `b`, `i`, `color`, `sz`, `u` — and a document that gets it wrong is one
Word calls corrupt, offers to repair, and repairs by stripping the formatting. Nothing catches it at
compile time. OpenXmlKit builds every properties element through the SDK's typed setters, which
place each child at its schema position, so the ordering is not something a caller can get wrong.

**Five unit scales, stringly typed.** Twips for page geometry, half-points for font size,
*eighths* of a point for border widths, EMUs for drawings, fiftieths of a percent for table widths —
variously `string`, `int`, `uint` and `StringValue`. A [`Length`](src/OpenXmlKit/Length.cs) carries a
distance and converts on the way out, so `Size = 12` is twelve points and a half-point border is
`Length.FromPoints(0.5)`.

**Toggle properties cannot be turned off.** `new Bold()` means on; there is no way to say
"explicitly not bold" against a bold paragraph style without knowing to write `w:val="0"`.
[`Toggle`](src/OpenXmlKit/Toggle.cs) has three states — on, off, and say-nothing — so a run inside a
bold style can be un-bolded, and an untouched font writes nothing at all.

**Built-in styles are absent from generated documents.** Word carries `TableGrid`, `Normal`,
`Heading1` and the rest at application level, and only writes them into a document when a user
inserts something that uses them. A document built in code names a style that is not there and
renders unstyled. `Styles.EnsureBuiltIn(BuiltInStyle.TableGrid)` writes Word's own definition,
brings the styles it depends on, and leaves an existing definition alone.


## Usage

Two ways to build, and they compose. The nested one suits a self-contained fragment:

<!-- snippet: NestedBuilder -->
<a id='snippet-NestedBuilder'></a>
```cs
using var document = Document.Create();

document.Body.AddTable(
    _ => _
        .Style("TableGrid")
        .Width(Width.Percent(100))
        .Row(
            row => row
                .Cell(Width.Percent(22), _ => _.AddParagraph(_ => _.Bold("Source")))
                .Cell(Width.Percent(78), "Budget paper 2")));

var bytes = document.ToArray();
```
<sup><a href='/src/OpenXmlKit.Tests/Samples.cs#L11-L26' title='Snippet source file'>snippet source</a> | <a href='#snippet-NestedBuilder' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The cursor suits a document written front to back:

<!-- snippet: CursorBuilder -->
<a id='snippet-CursorBuilder'></a>
```cs
using var document = Document.Create();
var builder = document.Builder;

builder.Heading(1, "Delivery update");
builder.Writeln("The commitment is on schedule.");

using (builder.PushFormatting())
{
    builder.Font.Bold = true;
    builder.Writeln("This paragraph is bold.");
}

using (builder.Table())
using (builder.Row())
{
    builder.InsertCell();
    builder.Write("A");
    builder.InsertCell();
    builder.Write("B");
}
```
<sup><a href='/src/OpenXmlKit.Tests/Samples.cs#L36-L59' title='Snippet source file'>snippet source</a> | <a href='#snippet-CursorBuilder' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`PushFormatting` scopes character, paragraph, cell, row and table formatting together, and every
paired start and end — table, row, bookmark — is a `using` block, so neither can be left unbalanced.


## Reading

Reading is a separate API, not the same one with the setters hidden. `DocumentView.Open` gives back
views — `ParagraphView`, `RunView`, `TableView` — which are lazy projections over the SDK tree and
have no way to change what they are looking at:

<!-- snippet: Reading -->
<a id='snippet-Reading'></a>
```cs
using var document = DocumentView.Open(source);

foreach (var paragraph in document.Body.Paragraphs)
{
    TestContext.Out.WriteLine(paragraph.Text);
}
```
<sup><a href='/src/OpenXmlKit.Tests/Samples.cs#L71-L80' title='Snippet source file'>snippet source</a> | <a href='#snippet-Reading' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The part worth knowing about is the resolver, which answers what formatting *applies* rather than
what is written:

<!-- snippet: ResolvingFormatting -->
<a id='snippet-ResolvingFormatting'></a>
```cs
var font = document.Formatting.FontFor(run, paragraph, tableStyleId: "Branded");
```
<sup><a href='/src/OpenXmlKit.Tests/Samples.cs#L97-L101' title='Snippet source file'>snippet source</a> | <a href='#snippet-ResolvingFormatting' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`FontFor` returns an `IFontView`, and it walks the cascade in the order the format defines — document defaults, table style, paragraph
style with its `basedOn` chain, character style, then direct formatting. That order is where most
Word surprises come from, and one in particular: **a paragraph style outranks a table style**, so
branding a table's font through its table style alone silently loses to whatever `Normal` says.


## Name collisions

OpenXmlKit gives its types the names they should have — `Paragraph`, `Table`, `Font`, `Style` — and
those are the names `DocumentFormat.OpenXml.Wordprocessing` already uses. A project importing both
gets an ambiguous reference on every one.

If your project does not import the SDK's namespace, there is nothing to do: `using OpenXmlKit.Word;`
and the names read as they should. If it does, opt into prefixed aliases instead:

```xml
<PropertyGroup>
  <OpenXmlKitAliases>W</OpenXmlKitAliases>
</PropertyGroup>
```

That gives `WParagraph`, `WTable`, `WFont` alongside the SDK's own names. Use `Word` for
`WordParagraph` instead. In alias mode, do **not** also import `OpenXmlKit.Word` — an alias avoids
the ambiguity by keeping the bare names out of scope, and importing the namespace puts them back.

The alias list is generated from the public API and checked by a test, so it cannot go stale, and
[`src/AliasCheck`](src/AliasCheck) is a project that compiles both sets of names in one file to prove
the mechanism works.


## Escape hatch

Every wrapper exposes the element underneath, and every container takes a raw element back:

<!-- snippet: EscapeHatch -->
<a id='snippet-EscapeHatch'></a>
```cs
// The SDK's own w:tbl, to hand to code that has not migrated yet.
var raw = table.ToOpenXml();

using var document = Document.Create();
document.Body.AppendElement(raw);
```
<sup><a href='/src/OpenXmlKit.Tests/Samples.cs#L111-L119' title='Snippet source file'>snippet source</a> | <a href='#snippet-EscapeHatch' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

This is deliberate and load-bearing. A library migrating onto OpenXmlKit does so a piece at a time,
and code that hands raw elements across a boundary has to keep working while it does.


## What v1 does not do

**Modify an existing document.** There are three ways in, and their types say what they do:

| | | |
|---|---|---|
| `Document.Create()` | `Document` | A new document. |
| `Document.OpenForAppend(...)` | `Document` | An existing one, to add content to — a branded template, typically, whose styles, headers and page setup the new content inherits. |
| `DocumentView.Open(...)` | `DocumentView` | An existing one, to read. |

What is missing is changing content that is already there, and it is missing at the type level
rather than by convention. A `Document` has nothing to enumerate — no `Body.Paragraphs`, no
`Table.Rows` — so there is no way to reach the content already in a file through the building API.
A `ParagraphView` has no `AddRun`, no `AddBookmark`, and its formatting is `IParagraphFormatView`,
which has no setters. So this does not compile, which is the whole point:

```cs
using var document = DocumentView.Open(bytes);
document.Body.Paragraphs.First().AddBookmark(...);   // no such method
document.Body.Paragraphs.First().Format.Alignment = ParagraphAlignment.Center;  // no setter
```

To inspect a document while building it, take a view of it: `DocumentView.Of(document)`.

**Charts, content controls, and OLE.** Reachable through the escape hatch, not modelled.

**Byte-identical packages.** Relationship ids are pinned, so the same calls produce byte-identical
part XML. The zip around them is not: its entries carry their own timestamps, written below the SDK.
If you need two runs to agree byte for byte — because you paginate documents yourself and write page
numbers in, say — add [DeterministicIoPackaging](https://github.com/SimonCropp/DeterministicIoPackaging),
which replaces that layer. It is deliberately your dependency rather than this package's, because it
changes how every document in the process is written.


## Verifying

```bash
dotnet build src --configuration Release -p:IsPackable=false
```

```bash
dotnet test src --configuration Release
```

Every test that produces a document runs it through `OpenXmlValidator`. That is the check the whole
emitter design exists to pass, and it is what would catch a schema-order regression — including one
introduced by an SDK upgrade, which
[`SchemaOrderTests`](src/OpenXmlKit.Tests/SchemaOrderTests.cs) pins directly.


## Icon

https://thenounproject.com/icon/phoenix-rising-6442478/
