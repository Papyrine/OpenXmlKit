namespace OpenXmlKit.Word;

public partial class ParagraphFormat
{
    IBordersView IParagraphFormatView.Borders => Borders;
    IShadingView IParagraphFormatView.Shading => Shading;
    IReadOnlyList<TabStop> IParagraphFormatView.TabStops => TabStops.AsList;
}
