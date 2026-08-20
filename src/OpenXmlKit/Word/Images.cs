using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXmlKit.Word;

/// <summary>
/// Builds the drawing element that puts a picture in a document.
/// </summary>
/// <remarks>
/// A single inline image is around ninety lines of nested DrawingML: a graphic wrapping a graphic
/// frame wrapping a picture, with the same extents restated at three levels, a preset geometry that
/// is always a rectangle, and unique ids that have to be allocated. None of it varies. Building it
/// here means an image is one call, and the EMU conversion never appears at a call site.
/// </remarks>
static class Images
{
    public static W.Drawing Build(
        string relationshipId,
        uint id,
        string name,
        Length width,
        Length height,
        ImageWrap wrap,
        string? description)
    {
        var extents = new A.Extents
        {
            Cx = width.Emu,
            Cy = height.Emu
        };

        var graphic = new A.Graphic(
            new A.GraphicData(
                new PIC.Picture(
                    new PIC.NonVisualPictureProperties(
                        new PIC.NonVisualDrawingProperties
                        {
                            Id = 0U,
                            Name = name,
                            Description = description
                        },
                        new PIC.NonVisualPictureDrawingProperties()),
                    new PIC.BlipFill(
                        new A.Blip
                        {
                            Embed = relationshipId
                        },
                        new A.Stretch(new A.FillRectangle())),
                    new PIC.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset
                            {
                                X = 0,
                                Y = 0
                            },
                            extents),
                        new A.PresetGeometry(new A.AdjustValueList())
                        {
                            Preset = A.ShapeTypeValues.Rectangle
                        })))
            {
                Uri = pictureNamespace
            });

        var docProperties = new DW.DocProperties
        {
            Id = id,
            Name = name,
            Description = description
        };

        if (wrap == ImageWrap.Inline)
        {
            return new(
                new DW.Inline(
                    new DW.Extent
                    {
                        Cx = width.Emu,
                        Cy = height.Emu
                    },
                    docProperties,
                    new DW.NonVisualGraphicFrameDrawingProperties(),
                    graphic)
                {
                    DistanceFromTop = 0U,
                    DistanceFromBottom = 0U,
                    DistanceFromLeft = 0U,
                    DistanceFromRight = 0U
                });
        }

        // A floated image is anchored rather than inline, and the text has to be told what to do
        // around it. Square wrapping on the matching side is what "float left" and "float right"
        // mean; without a wrap element the anchor renders behind the text.
        var anchor = new DW.Anchor(
            new DW.SimplePosition
            {
                X = 0,
                Y = 0
            },
            new DW.HorizontalPosition(
                new DW.HorizontalAlignment(wrap == ImageWrap.Right ? "right" : "left"))
            {
                RelativeFrom = DW.HorizontalRelativePositionValues.Margin
            },
            new DW.VerticalPosition(new DW.PositionOffset("0"))
            {
                RelativeFrom = DW.VerticalRelativePositionValues.Paragraph
            },
            new DW.Extent
            {
                Cx = width.Emu,
                Cy = height.Emu
            },
            new DW.EffectExtent
            {
                LeftEdge = 0,
                TopEdge = 0,
                RightEdge = 0,
                BottomEdge = 0
            },
            new DW.WrapSquare
            {
                WrapText = DW.WrapTextValues.BothSides
            },
            docProperties,
            new DW.NonVisualGraphicFrameDrawingProperties(),
            graphic)
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 114300U,
            DistanceFromRight = 114300U,
            SimplePos = false,
            RelativeHeight = 0U,
            BehindDoc = false,
            Locked = false,
            LayoutInCell = true,
            AllowOverlap = true
        };

        return new(anchor);
    }

    const string pictureNamespace = "http://schemas.openxmlformats.org/drawingml/2006/picture";

    public static PartTypeInfo PartTypeOf(ImageFormat format) =>
        format switch
        {
            ImageFormat.Jpeg => ImagePartType.Jpeg,
            ImageFormat.Gif => ImagePartType.Gif,
            ImageFormat.Bmp => ImagePartType.Bmp,
            ImageFormat.Tiff => ImagePartType.Tiff,
            _ => ImagePartType.Png
        };
}
