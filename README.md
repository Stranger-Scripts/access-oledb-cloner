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

## What it does NOT do (This is going to be implemented at a later date)

- Recreate non-PK indexes, constraints, foreign keys, triggers, views, stored procedures
- Perfectly replicate identity/autonumber semantics from SQL Server
- Guarantee BigInt support (mapped to `dbLong`)

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
