global using System.Diagnostics.CodeAnalysis;
global using System.Globalization;
global using System.Runtime.CompilerServices;
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
