// The non-generic collection interfaces, for the explicit IEnumerable.GetEnumerator every
// collection has to implement. Without this the short form binds to the generic IEnumerable<T>
// that the implicit usings supply, and the build breaks the moment an IDE cleanup simplifies the
// fully qualified name away.
global using System.Collections;
global using System.Diagnostics.CodeAnalysis;
global using System.Globalization;
global using System.Runtime.CompilerServices;
global using System.Text;
global using DocumentFormat.OpenXml;
global using DocumentFormat.OpenXml.Packaging;
global using OpenXmlKit;
global using OpenXmlKit.Word;

// The SDK's Wordprocessing namespace is reached through an alias rather than imported, because it
// owns Paragraph, Run, Table, Color, Font, Border, Shading and Style — the same names this library
// gives its own types. Aliasing keeps both in reach without either shadowing the other, and makes
// every line that touches raw OpenXML visibly do so.
global using W = DocumentFormat.OpenXml.Wordprocessing;
global using DocumentFormat.OpenXml.CustomProperties;
