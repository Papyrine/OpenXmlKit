# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this
repository.

## Project Overview

OpenXmlKit is an ergonomic wrapper over [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK)
for building and reading Word documents. It is intended as the shared low-level layer for the
Papyrine libraries — [OpenXmlHtml](https://github.com/Papyrine/OpenXmlHtml),
[Excelsior](https://github.com/Papyrine/Excelsior) and
[Parchment](https://github.com/Papyrine/Parchment) — each of which currently hand-rolls the same
OpenXML emission independently.

User-facing docs live in `readme.md`. Read it before making changes.

## Build & Test Commands

All commands run from the repository root.

```bash
dotnet build src --configuration Release -p:IsPackable=false
dotnet test src --configuration Release
dotnet test src/OpenXmlKit.Tests --filter "FullyQualifiedName~SchemaOrderTests"
```

Tests are NUnit under VSTest, so `dotnet test` with `--filter` is correct here. (Parchment uses
TUnit and needs `dotnet run --project` with `--treenode-filter` instead; that guidance is
Parchment-only.)

## Architecture

### The rule everything rests on

**Property containers are populated through the SDK's typed properties, never through `Append`.**

Word treats a properties element whose children are out of their `CT_*` schema sequence as a corrupt
document: it offers to repair on open, and repairing strips the formatting. Assigning through a
typed property (`properties.Bold = new()`) goes via `OpenXmlCompositeElement.SetElement`, which
places the child at its schema position; `Append` preserves insertion order and does not.

`SchemaOrderTests` pins this behaviour by assigning several containers in deliberately reversed
order and asserting the schema order comes out. If an SDK upgrade ever regresses it, that test fails
rather than the corruption reaching a document.

Two places genuinely cannot use typed properties, and both are commented as such:

- `Table.Flush` — `tblPr` and `tblGrid` are ordinary children of `w:tbl`, so their order is produced
  by removing and prepending in reverse. **`tblPr` is required by `CT_Tbl`, not optional**, so an
  empty one is written even when the table states no formatting; without it the validator's
  complaint lands on `tblGrid`, which is the next element rather than the missing one.
- `Section.Reference` — header and footer references are repeatable, and must lead `sectPr`.

### Two APIs, not one with the setters hidden

Building and reading are separate type hierarchies that share nothing but the enums and the
primitives.

- **Build** — `Document.Create` / `Document.OpenForAppend`, then `Body`, `Paragraph`, `Run`,
  `Table`, `Row`, `Cell`. Forward-only: content goes in, nothing comes back out. These types have
  no enumeration properties at all.
- **Read** — `DocumentView.Open` (or `DocumentView.Of` over a document being built), then
  `ParagraphView`, `RunView`, `TableView`, `RowView`, `CellView`, `SectionView`, `StyleView`, plus
  `HyperlinkView`, `ImageView`, `FieldView`, `FootnoteView` and `NumberingView` for the content
  that is not plain text. All `readonly struct`s over the SDK elements, allocating nothing but the
  enumerators, with no format caching and no flush machinery because reading needs neither.
  The rule the read side is held to: **anything the build API can write, the read API can read
  back.** `ReadContentTests` writes with one and asserts with the other so the two cannot drift.

Formatting crosses the line through interfaces: `IFontView`, `IParagraphFormatView` and the rest
carry the same properties as `Font` and friends with the setters removed, and the mutable classes
implement them. Sub-object properties need explicit implementation (`IShadingView IFontView.Shading
=> Shading;`) because C# does not accept a covariant return type for an implicit interface
implementation.

The reason for the split is that one type serving both jobs makes
`Open(...).Body.Paragraphs.First().AddBookmark(...)` compile, and it either does nothing or means
something the library does not do. Modifying existing content is out of scope for v1, and the
scope is now enforced by the compiler rather than by the readme. `OpenModeTests` asserts the
absence of the members that would break it.

### Deferred formatting, cascading flush

Content elements are appended to the tree as they are added. Formatting is deferred: each wrapper
holds a format object and rebuilds its properties element in `Flush()`. Every wrapper therefore owns
the wrappers it created (`Paragraph` owns its `Run`s, `Cell` its paragraphs, and so on) and
`Flush()` cascades down. `Document.Flush()` is the root, and is called by `Save`, `ToArray` and by
the read-side collection properties.

The reason: a caller sets formatting *after* creating a wrapper. An earlier design buffered content
instead and silently lost anything added after the wrapper was first inserted.

Flushing is idempotent — properties are rebuilt wholesale rather than patched — because reading an
element flushes it and a caller may keep editing afterwards.

### Layout

```
src/OpenXmlKit/
  Length.cs Toggle.cs Width.cs WidthUnit.cs   primitives, namespace OpenXmlKit
  Word/                                        everything else, namespace OpenXmlKit.Word
    Document.cs Document.Parts.cs              entry point, id allocation, side parts
    Body.cs BlockContainer.cs HeaderFooter.cs  block content containers
    Paragraph.cs Paragraph.Content.cs Run.cs   inline content; images, links, fields, footnotes
    Table.cs Row.cs Cell.cs                    tables
    Font.cs ParagraphFormat.cs CellFormat.cs   format objects (+ .Merge.cs partials)
    RowFormat.cs TableFormat.cs PageSetup.cs
    Border.cs Borders.cs Shading.cs TabStops.cs
    Styles.cs Style.cs BuiltInStyleDefinitions.cs
    Numbering.cs ListDefinition.cs ListLevel.cs
    TableStyleConditional.cs TableStyleArea.cs conditional table style formatting (tblStylePr)
    FormattingResolver.cs                      the read-side cascade
    Map.cs Toggles.cs WidthElement.cs Images.cs  internal converters
    Reading/                                   the read API: *View types and I*View interfaces
src/OpenXmlKit.Tests/    NUnit; DocumentAssert.IsValid runs OpenXmlValidator on everything
src/AliasCheck/          compile-only proof that alias mode works
src/OpenXmlKit.Benchmarks/
```

### Naming and the alias props

Public types are unprefixed (`Document`, `Paragraph`, `Table`, `Font`) in `OpenXmlKit.Word`. Those
are the same names `DocumentFormat.OpenXml.Wordprocessing` uses, so consumers importing both opt
into prefixed aliases by setting `OpenXmlKitAliases` to `W` or `Word`.

`src/OpenXmlKit/buildTransitive/OpenXmlKit.props` is **generated** — by `AliasPropsTests`, which
regenerates it and fails when it is out of date. Do not edit it by hand; add the type, run the
tests, commit the diff.

Inside this library the SDK is reached through the `W.` alias (`GlobalUsings.cs`) rather than
imported, for the same reason.

## Key Conventions

- Target frameworks `net48;net10.0` (matching OpenXmlHtml — consumers include net48). Tests are
  net10.0 only.
- `LangVersion` 14.0, `TreatWarningsAsErrors` and `EnforceCodeStyleInBuild` both on, so IDE style
  rules fail the build. Use `var` everywhere; `IDE0017` (object initialisers) and `IDE0021`
  (expression-bodied constructors) both fire as errors.
- **Text-parsing APIs are span-first**: the `ReadOnlySpan<char>` overload is the implementation and
  the `string` one delegates to it, returning the original instance when nothing changed. The
  consumers parse text they are reading rather than keeping — a CSS colour is a slice of a
  declaration, which is a slice of a style attribute — so a string-only signature puts a `ToString`
  on every boundary. `XmlChars` had to be retrofitted for this after Parchment could not consume
  it; `Color.TryParse` was done up front. Anything that follows them down here does the same.
- Central package management via `src/Directory.Packages.props`. Conventions otherwise come from the
  `ProjectDefaults` package, which also supplies the `Cancel`/`Date`/`Time` global aliases and
  copies `.editorconfig` into the repo.
- Relationship ids are pinned (`rStyles`, `rNumbering`, `rFootnotes`, `rImage{n}`) rather than
  generated, so the same calls produce byte-identical part XML. The *package* is still not
  byte-identical between runs — the zip entries carry their own timestamps, written below the SDK —
  and a consumer that needs that adds `DeterministicIoPackaging`, as the rest of the estate does.
  `DeterminismTests` pins both halves of this, including the limit.

## Non-obvious gotchas

- **`OpenXmlPackage.Clone` does not carry core properties.** Title, author and dates live outside
  the part graph it copies, and losing them is silent — the document opens fine with a blank
  properties dialog. `Document.CopyTo` puts them back by hand.
- **A `MemoryStream` built over a byte array is not expandable**, so `Document.OpenForAppend` over
  one fails on the first write with `NotSupportedException: Memory stream is not expandable`.
  `OpenForAppend(byte[])` exists so a caller holding bytes does not have to know that; the stream
  overload still has the hazard, and `OpenModeTests` pins both halves.
- **Characters XML 1.0 forbids do not fail until `Save`**, and the exception names none of the text
  that carried them. `XmlChars.Strip` is applied to every string the build API turns into a `w:t`,
  so a caller never has to. Three repos in the estate had their own copy of this — two of them
  character-for-character identical — which is why it is public rather than internal.
- **Editing existing content, when it arrives, cannot just flush a view.** Built wrappers rebuild
  their properties element wholesale, which is what makes flushing idempotent. Doing that to a
  paragraph read from a template would discard whatever the format model does not cover (`framePr`,
  a `sectPr` on the mark), so editing needs apply-in-place instead — and views would need owners.
- **A cell must contain a paragraph and must not end on a table**; both are repaired in
  `Cell.Flush`. `Cell.Container` deliberately skips that repair, because the cursor builder writes
  into a cell after opening it and the repair would leave an empty paragraph above everything.
- **`tblLook` is invalid inside a style's `tblPr`** (`CT_TblPrBase` has no such child), which is why
  `TableFormat.ToStyleProperties` exists alongside `ToProperties`.
- **A first-page or even-page header needs its section flag** (`titlePg`, `evenAndOddHeaders`) or
  Word stores it and never renders it. `Section.AddHeader` sets the flag rather than leaving the
  caller to discover the omission by its absence. The odd/even switch is document-scoped, not
  per-section, so it goes in the settings part.
- **Word bookmark names** must start with a letter, allow only letters, digits and underscores, and
  cap at 40 characters. A name that breaks the rules is dropped without a warning and every
  cross-reference to it renders as an error. `Bookmarks.Sanitise` handles the mechanical part, but
  sanitising collapses distinct titles onto the same name — derive names positionally where they
  have to be unique.
- **`ParagraphView.Text` walks `AllRuns`, not `Runs`.** A hyperlink holds its runs inside itself,
  so a reader that only looks at the paragraph's direct children drops the link text — silently,
  and all the way up through `DocumentView.Text`. `Runs` still means the direct children; only
  `Text` changed. `BlockContainerView.Text` includes tables for the same reason. `TextTests` pins
  both. The trailing empty paragraph Word requires after a table shows up as a blank line, and is
  reported rather than trimmed: a blank line a caller put there deliberately looks identical.
- **Cell margins are emitted as `w:start`/`w:end`, and that is correct.** Both pairs are in
  `CT_TcMar` and `CT_TblCellMar` — `start`/`end` are the Office 2010+ form, `left`/`right` the
  legacy one. Verified against Word 16.0: it reports identical padding for either, and on save it
  rewrites `start`/`end` to `left`/`right` with the values intact, which is proof it read them.
  Morph reads both and *prefers* `start`/`end`, with a test saying so because Excelsior emits them
  too. A comment in OpenXmlHtml's `PaddingHelper` claims `start`/`end` are "not schema-valid here
  and silently dropped by stricter consumers"; that is wrong on both counts. One real asymmetry:
  at table level `w:left` is `TableWidthDxaNilType` (integer, dxa or nil) where `w:start` is
  `TableWidthType`, so the legacy form cannot express a percentage — another reason to keep
  `start`/`end`.
- **A conditional table style block cannot state everything its format object can.** `tblStylePr`
  uses the `CT_*StyleOverride` types, which drop the properties that belong to content rather than
  to a style — `rStyle`, `tcW`, `gridSpan`, `tblW`, `tblLook` and the rest. `TableStyleConditional`
  throws on those rather than dropping them, because a dropped child is a style that silently does
  less than it says. The throw surfaces at flush, so it also comes out of `Dispose` if the caller
  never saved.
- **Fields need a cached value** or Word asks the reader for permission to update fields on open and
  shows a placeholder until they agree.
