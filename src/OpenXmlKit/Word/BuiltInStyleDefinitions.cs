namespace OpenXmlKit.Word;

/// <summary>
/// Word's own definitions of the built-in styles, as Word writes them.
/// </summary>
/// <remarks>
/// These are reproductions of what Word puts in styles.xml the first time a document uses each
/// style, not inventions. That matters: the point is that a document built in code renders the way
/// the same document authored in Word would, so a definition that merely looked reasonable would
/// defeat the exercise.
/// </remarks>
static class BuiltInStyleDefinitions
{
    public static string IdOf(BuiltInStyle style) =>
        style switch
        {
            BuiltInStyle.Normal => "Normal",
            BuiltInStyle.Heading1 => "Heading1",
            BuiltInStyle.Heading2 => "Heading2",
            BuiltInStyle.Heading3 => "Heading3",
            BuiltInStyle.Heading4 => "Heading4",
            BuiltInStyle.Heading5 => "Heading5",
            BuiltInStyle.Heading6 => "Heading6",
            BuiltInStyle.Title => "Title",
            BuiltInStyle.Subtitle => "Subtitle",
            BuiltInStyle.Quote => "Quote",
            BuiltInStyle.IntenseQuote => "IntenseQuote",
            BuiltInStyle.Caption => "Caption",
            BuiltInStyle.Header => "Header",
            BuiltInStyle.Footer => "Footer",
            BuiltInStyle.ListParagraph => "ListParagraph",
            BuiltInStyle.NoSpacing => "NoSpacing",
            BuiltInStyle.Hyperlink => "Hyperlink",
            BuiltInStyle.FootnoteText => "FootnoteText",
            BuiltInStyle.FootnoteReference => "FootnoteReference",
            BuiltInStyle.TableNormal => "TableNormal",
            BuiltInStyle.TableGrid => "TableGrid",
            _ => throw new ArgumentOutOfRangeException(nameof(style), style, "Unknown built-in style.")
        };

    /// <summary>
    /// The styles a definition is meaningless without, in the order they must be added.
    /// </summary>
    public static IEnumerable<BuiltInStyle> DependenciesOf(BuiltInStyle style) =>
        style switch
        {
            // TableGrid draws borders and nothing else; its cell padding comes from TableNormal
            // through basedOn, so the two are only useful together.
            BuiltInStyle.TableGrid => [BuiltInStyle.TableNormal],
            BuiltInStyle.Normal or BuiltInStyle.TableNormal => [],
            _ => [BuiltInStyle.Normal]
        };

    public static W.Style Build(BuiltInStyle style)
    {
        var element = style switch
        {
            BuiltInStyle.Normal => Normal(),
            BuiltInStyle.Heading1 => Heading(1, 16, "2F5496", 12, 0),
            BuiltInStyle.Heading2 => Heading(2, 13, "2F5496", 2, 0),
            BuiltInStyle.Heading3 => Heading(3, 12, "1F3763", 2, 0),
            BuiltInStyle.Heading4 => Heading(4, 11, "2F5496", 2, 0, italic: true),
            BuiltInStyle.Heading5 => Heading(5, 11, "2F5496", 2, 0),
            BuiltInStyle.Heading6 => Heading(6, 11, "1F3763", 2, 0),
            BuiltInStyle.Title => Title(),
            BuiltInStyle.Subtitle => Subtitle(),
            BuiltInStyle.Quote => Quote(),
            BuiltInStyle.IntenseQuote => IntenseQuote(),
            BuiltInStyle.Caption => Caption(),
            BuiltInStyle.Header => HeaderOrFooter("Header", "header"),
            BuiltInStyle.Footer => HeaderOrFooter("Footer", "footer"),
            BuiltInStyle.ListParagraph => ListParagraph(),
            BuiltInStyle.NoSpacing => NoSpacing(),
            BuiltInStyle.Hyperlink => Hyperlink(),
            BuiltInStyle.FootnoteText => FootnoteText(),
            BuiltInStyle.FootnoteReference => FootnoteReference(),
            BuiltInStyle.TableNormal => TableNormal(),
            BuiltInStyle.TableGrid => TableGrid(),
            _ => throw new ArgumentOutOfRangeException(nameof(style), style, "Unknown built-in style.")
        };

        return element;
    }

    static W.Style Paragraph(string id, string name) =>
        new()
        {
            Type = W.StyleValues.Paragraph,
            StyleId = id,
            StyleName = new()
            {
                Val = name
            }
        };

    static W.Style Normal()
    {
        var style = Paragraph("Normal", "Normal");
        style.Default = true;
        style.PrimaryStyle = new();
        return style;
    }

    static W.Style Heading(int level, double sizePoints, string color, double spaceBeforePoints, double spaceAfterPoints, bool italic = false)
    {
        var style = Paragraph($"Heading{level}", $"heading {level}");
        style.BasedOn = new()
        {
            Val = "Normal"
        };
        style.NextParagraphStyle = new()
        {
            Val = "Normal"
        };
        style.LinkedStyle = new()
        {
            Val = $"Heading{level}Char"
        };
        style.UIPriority = new()
        {
            Val = 9
        };
        style.PrimaryStyle = new();
        if (level > 1)
        {
            style.UnhideWhenUsed = new();
        }

        var paragraphFormat = new ParagraphFormat
        {
            SpaceBefore = Length.FromPoints(spaceBeforePoints),
            SpaceAfter = Length.FromPoints(spaceAfterPoints),
            KeepWithNext = Toggle.On,
            KeepTogether = Toggle.On,
            // The outline level is what puts a heading in the navigation pane and in a table of
            // contents. Word numbers it from zero, so Heading1 is level 0.
            OutlineLevel = level - 1
        };
        style.StyleParagraphProperties = Transfer<W.StyleParagraphProperties>(paragraphFormat.ToProperties());

        var font = new Font
        {
            NameAscii = "Calibri Light",
            NameHighAnsi = "Calibri Light",
            Size = Length.FromPoints(sizePoints),
            Color = Color.Parse(color)
        };
        if (italic)
        {
            font.Italic = Toggle.On;
        }

        style.StyleRunProperties = Transfer<W.StyleRunProperties>(font.ToProperties());
        return style;
    }

    static W.Style Title()
    {
        var style = Paragraph("Title", "Title");
        style.BasedOn = new()
        {
            Val = "Normal"
        };
        style.NextParagraphStyle = new()
        {
            Val = "Normal"
        };
        style.UIPriority = new()
        {
            Val = 10
        };
        style.PrimaryStyle = new();
        style.StyleParagraphProperties = Transfer<W.StyleParagraphProperties>(
            new ParagraphFormat
            {
                SpaceAfter = Length.FromPoints(4),
                ContextualSpacing = Toggle.On
            }.ToProperties());
        style.StyleRunProperties = Transfer<W.StyleRunProperties>(
            new Font
            {
                NameAscii = "Calibri Light",
                NameHighAnsi = "Calibri Light",
                Size = Length.FromPoints(28),
                CharacterSpacing = Length.FromTwips(-10)
            }.ToProperties());
        return style;
    }

    static W.Style Subtitle()
    {
        var style = Paragraph("Subtitle", "Subtitle");
        style.BasedOn = new()
        {
            Val = "Normal"
        };
        style.NextParagraphStyle = new()
        {
            Val = "Normal"
        };
        style.UIPriority = new()
        {
            Val = 11
        };
        style.PrimaryStyle = new();
        style.StyleParagraphProperties = Transfer<W.StyleParagraphProperties>(
            new ParagraphFormat
            {
                SpaceAfter = Length.FromPoints(8)
            }.ToProperties());
        style.StyleRunProperties = Transfer<W.StyleRunProperties>(
            new Font
            {
                Size = Length.FromPoints(14),
                Color = Color.Parse("5A5A5A"),
                CharacterSpacing = Length.FromTwips(15)
            }.ToProperties());
        return style;
    }

    static W.Style Quote()
    {
        var style = Paragraph("Quote", "Quote");
        style.BasedOn = new()
        {
            Val = "Normal"
        };
        style.NextParagraphStyle = new()
        {
            Val = "Normal"
        };
        style.UIPriority = new()
        {
            Val = 29
        };
        style.PrimaryStyle = new();
        style.StyleParagraphProperties = Transfer<W.StyleParagraphProperties>(
            new ParagraphFormat
            {
                SpaceBefore = Length.FromPoints(8),
                SpaceAfter = Length.FromPoints(8),
                Alignment = ParagraphAlignment.Center
            }.ToProperties());
        style.StyleRunProperties = Transfer<W.StyleRunProperties>(
            new Font
            {
                Italic = Toggle.On,
                Color = Color.Parse("404040")
            }.ToProperties());
        return style;
    }

    static W.Style IntenseQuote()
    {
        var style = Paragraph("IntenseQuote", "Intense Quote");
        style.BasedOn = new()
        {
            Val = "Normal"
        };
        style.NextParagraphStyle = new()
        {
            Val = "Normal"
        };
        style.UIPriority = new()
        {
            Val = 30
        };
        style.PrimaryStyle = new();

        var format = new ParagraphFormat
        {
            SpaceBefore = Length.FromPoints(18),
            SpaceAfter = Length.FromPoints(18),
            LeftIndent = Length.FromInches(0.6),
            RightIndent = Length.FromInches(0.6),
            Alignment = ParagraphAlignment.Center
        };
        format.Borders.Top.Set(BorderStyle.Single, Length.FromEighthPoints(4), Color.Parse("4472C4"));
        format.Borders.Bottom.Set(BorderStyle.Single, Length.FromEighthPoints(4), Color.Parse("4472C4"));
        style.StyleParagraphProperties = Transfer<W.StyleParagraphProperties>(format.ToProperties());

        style.StyleRunProperties = Transfer<W.StyleRunProperties>(
            new Font
            {
                Italic = Toggle.On,
                Color = Color.Parse("4472C4")
            }.ToProperties());
        return style;
    }

    static W.Style Caption()
    {
        var style = Paragraph("Caption", "caption");
        style.BasedOn = new()
        {
            Val = "Normal"
        };
        style.NextParagraphStyle = new()
        {
            Val = "Normal"
        };
        style.UIPriority = new()
        {
            Val = 35
        };
        style.SemiHidden = new();
        style.UnhideWhenUsed = new();
        style.PrimaryStyle = new();
        style.StyleParagraphProperties = Transfer<W.StyleParagraphProperties>(
            new ParagraphFormat
            {
                SpaceAfter = Length.FromPoints(10),
                LineSpacingMultiple = 1
            }.ToProperties());
        style.StyleRunProperties = Transfer<W.StyleRunProperties>(
            new Font
            {
                Italic = Toggle.On,
                Size = Length.FromPoints(9),
                Color = Color.Parse("44546A")
            }.ToProperties());
        return style;
    }

    static W.Style HeaderOrFooter(string id, string name)
    {
        var style = Paragraph(id, name);
        style.BasedOn = new()
        {
            Val = "Normal"
        };
        style.LinkedStyle = new()
        {
            Val = id + "Char"
        };
        style.UIPriority = new()
        {
            Val = 99
        };
        style.UnhideWhenUsed = new();

        // The centre and right tab stops are the whole point of these styles: they are what put the
        // three usual header slots where a reader expects them, on a default A4 or Letter page.
        var format = new ParagraphFormat
        {
            SpaceAfter = Length.Zero,
            LineSpacingMultiple = 1
        };
        format.TabStops.Add(Length.FromInches(3.25), TabAlignment.Center);
        format.TabStops.Add(Length.FromInches(6.5), TabAlignment.Right);
        style.StyleParagraphProperties = Transfer<W.StyleParagraphProperties>(format.ToProperties());
        return style;
    }

    static W.Style ListParagraph()
    {
        var style = Paragraph("ListParagraph", "List Paragraph");
        style.BasedOn = new()
        {
            Val = "Normal"
        };
        style.UIPriority = new()
        {
            Val = 34
        };
        style.PrimaryStyle = new();
        style.StyleParagraphProperties = Transfer<W.StyleParagraphProperties>(
            new ParagraphFormat
            {
                LeftIndent = Length.FromInches(0.5),
                // Without this a bulleted list is spaced as though every item were its own
                // paragraph, which is what makes a hand-built list look wrong.
                ContextualSpacing = Toggle.On
            }.ToProperties());
        return style;
    }

    static W.Style NoSpacing()
    {
        var style = Paragraph("NoSpacing", "No Spacing");
        style.UIPriority = new()
        {
            Val = 1
        };
        style.PrimaryStyle = new();
        style.StyleParagraphProperties = Transfer<W.StyleParagraphProperties>(
            new ParagraphFormat
            {
                SpaceAfter = Length.Zero,
                LineSpacingMultiple = 1
            }.ToProperties());
        return style;
    }

    static W.Style Hyperlink()
    {
        var font = new Font
        {
            Color = Color.Parse("0563C1"),
            Underline = UnderlineStyle.Single
        };
        return new()
        {
            Type = W.StyleValues.Character,
            StyleId = "Hyperlink",
            StyleName = new()
            {
                Val = "Hyperlink"
            },
            BasedOn = new()
            {
                Val = "DefaultParagraphFont"
            },
            UIPriority = new()
            {
                Val = 99
            },
            UnhideWhenUsed = new(),
            StyleRunProperties = Transfer<W.StyleRunProperties>(font.ToProperties())
        };
    }

    static W.Style FootnoteText()
    {
        var style = Paragraph("FootnoteText", "footnote text");
        style.BasedOn = new()
        {
            Val = "Normal"
        };
        style.LinkedStyle = new()
        {
            Val = "FootnoteTextChar"
        };
        style.UIPriority = new()
        {
            Val = 99
        };
        style.SemiHidden = new();
        style.UnhideWhenUsed = new();
        style.StyleParagraphProperties = Transfer<W.StyleParagraphProperties>(
            new ParagraphFormat
            {
                SpaceAfter = Length.Zero,
                LineSpacingMultiple = 1
            }.ToProperties());
        style.StyleRunProperties = Transfer<W.StyleRunProperties>(
            new Font
            {
                Size = Length.FromPoints(10)
            }.ToProperties());
        return style;
    }

    static W.Style FootnoteReference()
    {
        var font = new Font
        {
            VerticalPosition = VerticalTextPosition.Superscript
        };
        return new()
        {
            Type = W.StyleValues.Character,
            StyleId = "FootnoteReference",
            StyleName = new()
            {
                Val = "footnote reference"
            },
            BasedOn = new()
            {
                Val = "DefaultParagraphFont"
            },
            UIPriority = new()
            {
                Val = 99
            },
            SemiHidden = new(),
            UnhideWhenUsed = new(),
            StyleRunProperties = Transfer<W.StyleRunProperties>(font.ToProperties())
        };
    }

    // Mirrors the TableNormal Word ships in user-authored docs: zero indentation, 108 twip
    // left/right cell margins, marked default so a table that names no style still gets the
    // padding.
    static W.Style TableNormal()
    {
        var style = new W.Style
        {
            Type = W.StyleValues.Table,
            StyleId = "TableNormal",
            Default = true,
            StyleName = new()
            {
                Val = "Normal Table"
            },
            UIPriority = new()
            {
                Val = 99
            },
            SemiHidden = new(),
            UnhideWhenUsed = new()
        };

        var format = new TableFormat
        {
            Indent = Length.Zero
        };
        format.SetDefaultMargins(Length.FromTwips(108), Length.Zero);
        style.Append(format.ToStyleProperties());
        return style;
    }

    // Mirrors the TableGrid Word ships when a table is inserted from the ribbon: borders only, no
    // cell margins of its own. Padding flows in through basedOn.
    static W.Style TableGrid()
    {
        var paragraphFormat = new ParagraphFormat
        {
            SpaceAfter = Length.Zero,
            LineSpacingMultiple = 1
        };
        var style = new W.Style
        {
            Type = W.StyleValues.Table,
            StyleId = "TableGrid",
            StyleName = new()
            {
                Val = "Table Grid"
            },
            BasedOn = new()
            {
                Val = "TableNormal"
            },
            UIPriority = new()
            {
                Val = 39
            },
            StyleParagraphProperties = Transfer<W.StyleParagraphProperties>(paragraphFormat.ToProperties())
        };

        var format = new TableFormat();
        format.Borders.SetAll(BorderStyle.Single, Length.FromEighthPoints(4), Color.Auto);
        foreach (var border in format.Borders.All())
        {
            border.Space = Length.Zero;
        }

        style.Append(format.ToStyleProperties());
        return style;
    }

    // The style-scoped property elements hold the same children in the same order as the
    // content-scoped ones, so the ordering the SDK applied when the source was built carries over
    // rather than being restated here.
    static T Transfer<T>(OpenXmlElement? source)
        where T : OpenXmlElement, new()
    {
        var target = new T();
        if (source == null)
        {
            return target;
        }

        foreach (var child in source.ChildElements.ToList())
        {
            child.Remove();
            target.AppendChild(child);
        }

        return target;
    }
}
