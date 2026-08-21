# Schema data

Copied unmodified from [dotnet/Open-XML-SDK](https://github.com/dotnet/Open-XML-SDK):

| here | there |
|---|---|
| `wordprocessingml.json` | `data/schemas/schemas_openxmlformats_org_wordprocessingml_2006_main.json` |
| `namespaces.json` | `data/namespaces.json` |

This is the data the SDK's own code generator reads particle order from, which is why
`SchemaOrderTable` reads it too rather than anyone transcribing the sequences by hand.

Unmodified on purpose. Trimming it to the four fields the generator reads would take it from 1.8 MB
to 85 KB, and would put a hand step between the SDK's data and this repo that nothing tests — the
same reason the extraction is a test rather than a script. Updating for a newer SDK means copying
both files across again and running the suite, which rewrites `SchemaOrder.Table.cs` and fails so
the diff gets read.
