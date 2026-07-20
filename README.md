# access-oledb-cloner

# Access OLE DB Clone (VBA)

Clone tables (schema + primary keys + data) from an OLE DB source into a local Microsoft Access database.

This exists because Access can be painful/limited when you want to pull full table structures + data from external
sources through OLE DB in a repeatable way.

## What it does

- Enumerates remote tables via `OpenSchema(adSchemaTables)`
- Creates matching local tables (DAO TableDefs/Fields)
- Attempts to create primary keys (via `adSchemaPrimaryKeys`)
- Copies all rows into local tables in batched transactions
- Matches source columns to local columns **by name**, so re-running against an
  existing local table with a different column order stays correct
- Continues past a table it cannot clone, then reports every failure at the end

## Planned

- Recreate non-PK indexes (via `adSchemaIndexes`)
- Recreate foreign keys as DAO `Relations` (needs dependency-ordered creation;
  will not handle cyclic or self-referencing schemas)

## What it will not do

These are limitations of Access itself, not gaps in the script:

- **Triggers, stored procedures, views.** Access has no equivalent for the first
  two; view definitions are T-SQL and do not translate to Access SQL.
- **CHECK constraints.** Access's support is too inconsistent to rely on.
- **Identity/autonumber semantics.** Access reassigns autonumber values on
  insert, so preserving the *source* values means storing them as plain numbers.
  That is what this script does, and it is the correct trade-off for a clone —
  the values match, but the local column will not auto-increment.

## Type mapping notes

- **64-bit integers** (`adBigInt`, `adUnsignedInt`, `adUnsignedBigInt`) map to
  `dbBigInt` where the Access engine supports it. Support is probed once at
  runtime; where unavailable they fall back to `dbDouble`, which covers the full
  range but is only exact up to 2^53. (They are *not* mapped to `dbLong`, which
  would silently overflow.)
- **GUIDs** are stored as 38-char text — the width of the string form, not the
  16-byte `DefinedSize` the provider reports.
- **`adDecimal` / `adNumeric`** map to `dbDouble` and can lose precision on
  high-precision decimal columns. Known limitation.

## Requirements

- Microsoft Access (VBA)
- An installed OLE DB provider suitable for your source (ACE, SQL Server, etc.)
- Permissions to read schema + data from the source

## Install

1. Download `src/modOleDbClone.bas`
2. Open your Access `.accdb`
3. VBA editor → **File → Import File…** → import the `.bas` module

## Usage

### Minimal

```vb
Public Sub RunClone()
    Dim conn As String
    conn = "Provider=...;Data Source=...;User ID=...;Password=...;"

    CloneOleDbToAccess_Default conn
End Sub
```

Note that `CloneOleDbToAccess_Default` is the opinionated preset (replace tables,
batch 500, verbose, skip system tables). If you pass your own `CloneOptions`, the
Boolean fields default to `False` — an unset VBA `Boolean` is indistinguishable
from an explicit `False`, so set the ones you want. See
[`example/example_runclone.bas`](example/example_runclone.bas).

## License

[MIT](LICENSE) © Gabriel Geissler
