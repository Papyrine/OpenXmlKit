# OpenXmlKit

An ergonomic wrapper over the OpenXML SDK for building and reading Word documents.

Building and reading are separate APIs rather than one with the setters hidden, so content read
from a document cannot be assigned to by mistake. Property containers are populated through the
SDK's typed properties, which is what keeps their children in the schema order Word treats as the
difference between a document and a corrupt one.

Covers paragraphs and runs, tables, styles including the conditional blocks a table style is mostly
made of, numbering, sections, headers and footers, images, hyperlinks, fields, footnotes and
document properties — and handles the characters XML forbids, which otherwise surface as an
exception at save time naming none of the text that carried them.

See the [readme](https://github.com/Papyrine/OpenXmlKit) for usage.
