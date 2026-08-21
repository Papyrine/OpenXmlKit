// The guarantee the whole emitter layer is built on.
//
// Word treats a properties element whose children are out of their CT_* schema sequence as a
// corrupt document: it offers to "repair" it on open, and repairing strips the formatting. The
// estate deals with this by hand — Excelsior appends in a memorised order, OpenXmlHtml appends in
// a convenient order and re-sorts afterwards (ReorderCellProperties) — and both carry comments
// explaining the sequence to the next reader.
//
// None of that is necessary, because the SDK already knows every sequence: assigning through a
// typed property goes via OpenXmlCompositeElement.SetElement, which places the child at its
// schema position. Append does not, which is why Append-based code has to know the order.
//
// So OpenXmlKit's rule is: property containers are populated through typed properties, never
// through Append. These tests are what makes that rule safe to rely on — if an SDK upgrade ever
// regresses the behaviour, this fails rather than shipping documents Word calls corrupt.
#pragma warning disable IDE0017 // assignment order is the subject of these tests
[TestFixture]
public class SchemaOrderTests
{
    [Test]
    public void RunProperties_AssignedInReverse_EmitInSchemaOrder()
    {
        var properties = new W.RunProperties
        {
            VerticalTextAlignment = new()
            {
                Val = W.VerticalPositionValues.Superscript
            },
            Shading = new()
            {
                Val = W.ShadingPatternValues.Clear,
                Fill = "FFFF00"
            },
            Underline = new()
            {
                Val = W.UnderlineValues.Single
            },
            FontSize = new()
            {
                Val = "24"
            },
            Color = new()
            {
                Val = "FF0000"
            },
            Strike = new(),
            SmallCaps = new(),
            Italic = new(),
            Bold = new(),
            RunFonts = new()
            {
                Ascii = "Calibri"
            }
        };

        Assert.That(
            LocalNames(properties),
            Is.EqualTo(["rFonts", "b", "i", "smallCaps", "strike", "color", "sz", "u", "shd", "vertAlign"]));
    }

    [Test]
    public void TableCellProperties_AssignedInReverse_EmitInSchemaOrder()
    {
        var properties = new W.TableCellProperties
        {
            TableCellVerticalAlignment = new()
            {
                Val = W.TableVerticalAlignmentValues.Center
            },
            TableCellMargin = new(),
            Shading = new()
            {
                Val = W.ShadingPatternValues.Clear,
                Fill = "E0E8F2"
            },
            TableCellBorders = new(),
            VerticalMerge = new()
            {
                Val = W.MergedCellValues.Restart
            },
            GridSpan = new()
            {
                Val = 2
            },
            TableCellWidth = new()
            {
                Width = "5000",
                Type = W.TableWidthUnitValues.Pct
            }
        };

        Assert.That(
            LocalNames(properties),
            Is.EqualTo(["tcW", "gridSpan", "vMerge", "tcBorders", "shd", "tcMar", "vAlign"]));
    }

    [Test]
    public void ParagraphProperties_AssignedInReverse_EmitInSchemaOrder()
    {
        var properties = new W.ParagraphProperties
        {
            Justification = new()
            {
                Val = W.JustificationValues.Center
            },
            Indentation = new()
            {
                Left = "720"
            },
            SpacingBetweenLines = new()
            {
                After = "120"
            },
            Shading = new()
            {
                Val = W.ShadingPatternValues.Clear,
                Fill = "EEEEEE"
            },
            ParagraphBorders = new(),
            KeepNext = new(),
            ParagraphStyleId = new()
            {
                Val = "Heading1"
            }
        };

        Assert.That(
            LocalNames(properties),
            Is.EqualTo([
                "pStyle",
                "keepNext",
                "pBdr",
                "shd",
                "spacing",
                "ind",
                "jc"
            ]));
    }

    [Test]
    public void TableProperties_AssignedInReverse_EmitInSchemaOrder()
    {
        var properties = new W.TableProperties
        {
            TableLook = new()
            {
                Val = "04A0"
            },
            TableCellMarginDefault = new(),
            TableLayout = new()
            {
                Type = W.TableLayoutValues.Fixed
            },
            TableBorders = new(),
            TableWidth = new()
            {
                Width = "5000",
                Type = W.TableWidthUnitValues.Pct
            },
            TableStyle = new()
            {
                Val = "TableGrid"
            }
        };

        Assert.That(
            LocalNames(properties),
            Is.EqualTo([
                "tblStyle",
                "tblW",
                "tblBorders",
                "tblLayout",
                "tblCellMar",
                "tblLook"
            ]));
    }

    [Test]
    public void Style_AssignedInReverse_EmitInSchemaOrder()
    {
        var style = new W.Style
        {
            StyleRunProperties = new(),
            StyleParagraphProperties = new(),
            PrimaryStyle = new(),
            UIPriority = new()
            {
                Val = 39
            },
            NextParagraphStyle = new()
            {
                Val = "Normal"
            },
            BasedOn = new()
            {
                Val = "Normal"
            },
            StyleName = new()
            {
                Val = "Heading 1"
            }
        };

        Assert.That(
            LocalNames(style),
            Is.EqualTo([
                "name",
                "basedOn",
                "next",
                "uiPriority",
                "qFormat",
                "pPr",
                "rPr"
            ]));
    }

    [Test]
    public void Append_PreservesInsertionOrder_WhichIsWhyItIsNotUsed()
    {
        // The counterexample. This is the behaviour that forces every Append-based caller in the
        // estate to know the sequence; it is documented here so the rule above has a reason
        // attached rather than being folklore.
        var properties = new W.RunProperties();
        properties.Append(
            new W.VerticalTextAlignment
            {
                Val = W.VerticalPositionValues.Superscript
            });
        properties.Append(
            new W.Underline
            {
                Val = W.UnderlineValues.Single
            });
        properties.Append(new W.Bold());
        properties.Append(
            new W.RunFonts
            {
                Ascii = "Calibri"
            });

        Assert.That(
            LocalNames(properties),
            Is.EqualTo([
                "vertAlign",
                "u",
                "b",
                "rFonts"
            ]));
    }

    [Test]
    public void SectionPropertiesArePlacedRatherThanAppended()
    {
        // sectPr is the one properties container with no typed properties to assign through:
        // header and footer references repeat, so the SDK generates none. Appending is therefore
        // the only option the SDK offers, and appending is wrong as soon as sectPr holds a child
        // this library does not model.
        using var stream = new MemoryStream();
        using (var document = Document.Create(stream))
        {
            document.Body.Paragraph("text");
            var section = document.Body.Section;

            section.ToOpenXml()
                .Append(
                    new W.PageBorders(
                        new W.TopBorder
                        {
                            Val = W.BorderValues.Single,
                            Size = 4
                        }));

            section.PageSetup.PageWidth = Length.FromInches(8.5);
            section.PageSetup.PageHeight = Length.FromInches(11);
        }

        DocumentAssert.IsValid(stream.ToArray());

        using var read = DocumentView.Open(stream.ToArray());
        var properties = read.MainPart.Document!.Body!.GetFirstChild<W.SectionProperties>()!;

        // pgSz precedes pgBorders in CT_SectPr, and did not before this was pinned.
        Assert.That(LocalNames(properties), Is.EqualTo(["pgSz", "pgBorders"]));
    }

    [Test]
    public void HeaderReferencesStillLeadTheSection()
    {
        using var document = Document.Create();
        var section = document.Body.Section;
        section.PageSetup.PageWidth = Length.FromInches(8.5);
        section.AddHeader().AddParagraph("header");

        var bytes = DocumentAssert.IsValid(document);

        using var read = DocumentView.Open(bytes);
        var properties = read.MainPart.Document!.Body!.GetFirstChild<W.SectionProperties>()!;
        Assert.That(LocalNames(properties).First(), Is.EqualTo("headerReference"));
    }

    [Test]
    public void TheTableAgreesWithTheSequencesThisLibraryReliesOn()
    {
        // Spot checks written from the specification rather than from the generator, so a bad
        // regeneration is caught by something that read the order independently.
        Assert.That(
            SchemaOrder.IndexOf(new W.SectionProperties(), new W.PageSize()),
            Is.LessThan(SchemaOrder.IndexOf(new W.SectionProperties(), new W.PageBorders())));

        Assert.That(
            SchemaOrder.IndexOf(new W.TableCellProperties(), new W.GridSpan()),
            Is.LessThan(SchemaOrder.IndexOf(new W.TableCellProperties(), new W.VerticalMerge())));

        // The case the SDK's own typed properties cannot reach: documentProtection is the
        // thirty-fifth child of CT_Settings, and the generator stops emitting them at twenty-one.
        Assert.That(
            SchemaOrder.IndexOf(new W.Settings(), new W.DocumentProtection()),
            Is.GreaterThan(SchemaOrder.IndexOf(new W.Settings(), new W.Zoom())));
        Assert.That(
            SchemaOrder.IndexOf(new W.Settings(), new W.DocumentProtection()),
            Is.LessThan(SchemaOrder.IndexOf(new W.Settings(), new W.Compatibility())));
    }

    [Test]
    public void PlaceInsertsIntoALongSequenceTheSdkGivesNoTypedPropertyFor()
    {
        // What a caller hand-maintains a list of thirty-four types to do today.
        // In the sequence the schema states, as a real settings part would be.
        var settings = new W.Settings();
        settings.AppendChild(new W.DefaultTabStop { Val = 720 });
        settings.AppendChild(new W.Compatibility());
        settings.AppendChild(new W.Rsids());

        SchemaOrder.Place(
            settings,
            new W.DocumentProtection
            {
                Edit = W.DocumentProtectionValues.ReadOnly,
                Enforcement = true
            });

        Assert.That(
            LocalNames(settings),
            Is.EqualTo(["documentProtection", "defaultTabStop", "compat", "rsids"]));
    }

    [Test]
    public void PlaceReplacesRatherThanDuplicating()
    {
        var properties = new W.SectionProperties();
        SchemaOrder.Place(properties, new W.PageSize { Width = 100U });
        SchemaOrder.Place(properties, new W.PageSize { Width = 200U });

        Assert.That(LocalNames(properties), Is.EqualTo(["pgSz"]));
        Assert.That(properties.GetFirstChild<W.PageSize>()!.Width!.Value, Is.EqualTo(200U));
    }

    [Test]
    public void AChildTheSequenceDoesNotMentionIsAppended()
    {
        // A choice imposes no order, so there is nothing to place against and nothing lost by
        // appending — which is what the SDK would have done anyway.
        var properties = new W.SectionProperties();
        properties.AppendChild(new W.PageSize());
        SchemaOrder.Place(properties, new W.Bold());

        Assert.That(LocalNames(properties), Is.EqualTo(["pgSz", "b"]));
    }

    static List<string> LocalNames(OpenXmlElement element) =>
        element.ChildElements.Select(_ => _.LocalName).ToList();
}
