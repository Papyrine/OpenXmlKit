global using System.Globalization;
global using System.Text.RegularExpressions;
global using System.Xml.Linq;
global using DocumentFormat.OpenXml;
global using DocumentFormat.OpenXml.Packaging;
global using NUnit.Framework;
global using OpenXmlKit;
global using OpenXmlKit.Word;

// The same collision consumers hit, handled the same way. OpenXmlKit.Word and the SDK's
// Wordprocessing namespace both own Document, Paragraph, Run, Table, Color, Font and Style, so the
// SDK is reached through an alias rather than imported. Consumers that would rather keep the bare
// SDK names can set OpenXmlKitAliases in their project instead and get prefixed aliases for these
// types — see buildTransitive/OpenXmlKit.props.
global using W = DocumentFormat.OpenXml.Wordprocessing;
global using System.Text;
