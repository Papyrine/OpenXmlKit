namespace OpenXmlKit.Word;

public enum TableLayout
{
    /// <summary>
    /// Word fits columns to their content, treating declared widths as hints.
    /// </summary>
    Autofit,

    /// <summary>
    /// Declared column widths are honoured as written.
    /// </summary>
    Fixed
}
