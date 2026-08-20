// A consumer that already lives in the SDK's Wordprocessing namespace, as every Papyrine library
// that touches Word does. Note what is absent: OpenXmlKit.Word is not imported. That is the point
// of alias mode — the bare names never enter scope, so Paragraph keeps meaning the SDK's and
// WParagraph means this library's.
global using DocumentFormat.OpenXml;
global using DocumentFormat.OpenXml.Packaging;
global using DocumentFormat.OpenXml.Wordprocessing;
