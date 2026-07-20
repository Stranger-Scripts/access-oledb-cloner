Option Compare Database
Option Explicit

' =============================================================================
' MS Access: Clone OLE DB database schema + data into local Access tables
'
' - Creates local tables based on remote schema
' - Creates primary keys where possible
' - Copies data row-by-row in transactions (batched)
'
' Notes:
' - Uses late-bound ADODB to reduce reference issues.
' - DAO is required (native in Access).
' =============================================================================

' ---------------------------
' CONSTANTS
' ---------------------------

' dbBigInt is only defined in newer DAO type libraries. Referencing the named
' constant would break compilation on older Access installs, so use the literal.
Private Const DAO_dbBigInt As Long = 16

' Length of a GUID in its string form: {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}
Private Const GUID_TEXT_SIZE As Long = 38

Private mBigIntProbed As Boolean
Private mBigIntSupported As Boolean

' ---------------------------
' PUBLIC API
' ---------------------------

Public Type CloneOptions
    ReplaceLocalTables As Boolean   ' Drop and recreate local tables
    BatchSize As Long               ' Commit every N rows
    Verbose As Boolean              ' Debug.Print progress

    ' Table filters (simple prefix filters)
    SkipSystemTables As Boolean     ' Skips MSys*, sys_*, dt_* by default
End Type

Public Sub CloneOleDbToAccess( _
    ByVal oleDbConnString As String, _
    Optional ByVal options As CloneOptions)

    Dim cn As Object        ' ADODB.Connection (late bound)
    Dim rsTables As Object  ' ADODB.Recordset
    Dim tblName As String
    Dim failures As String
    Dim failCount As Long

    ApplyDefaultOptions options

    Set cn = CreateObject("ADODB.Connection")
    cn.Open oleDbConnString

    If options.Verbose Then Debug.Print "=== Creating schema (tables + PKs) ==="

    Set rsTables = cn.OpenSchema(20, Array(Empty, Empty, Empty, "TABLE")) ' adSchemaTables=20
    Do While Not rsTables.EOF
        tblName = Nz(rsTables.Fields("TABLE_NAME").Value, vbNullString)

        If ShouldProcessTable(tblName, options) Then
            If options.Verbose Then Debug.Print "Schema for: "; tblName

            ' A single unclonable table must not abort the whole run.
            On Error Resume Next
            CreateTableWithSchema cn, tblName, options
            If Err.Number <> 0 Then
                failCount = failCount + 1
                failures = failures & vbCrLf & "  schema [" & tblName & "]: " & Err.Description
                Debug.Print "SCHEMA FAILED for [" & tblName & "]: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If

        rsTables.MoveNext
    Loop
    rsTables.Close

    If options.Verbose Then Debug.Print "=== Copying data ==="

    Set rsTables = cn.OpenSchema(20, Array(Empty, Empty, Empty, "TABLE"))
    Do While Not rsTables.EOF
        tblName = Nz(rsTables.Fields("TABLE_NAME").Value, vbNullString)

        If ShouldProcessTable(tblName, options) Then
            If TableExists(tblName) Then
                If options.Verbose Then Debug.Print "Copying data for: "; tblName

                On Error Resume Next
                BulkCopyTableData cn, tblName, options
                If Err.Number <> 0 Then
                    failCount = failCount + 1
                    failures = failures & vbCrLf & "  data [" & tblName & "]: " & Err.Description
                    Debug.Print "DATA COPY FAILED for [" & tblName & "]: " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If

        rsTables.MoveNext
    Loop
    rsTables.Close

    cn.Close
    Set cn = Nothing

    If failCount = 0 Then
        MsgBox "Done cloning DB (schema + data).", vbInformation
    Else
        MsgBox "Clone finished with " & failCount & " failure(s):" & vbCrLf & failures & _
               vbCrLf & vbCrLf & "See the Immediate window for details.", vbExclamation
    End If
End Sub

' Convenience wrapper for people who like “one-liner” entry points
Public Sub CloneOleDbToAccess_Default(ByVal oleDbConnString As String)
    Dim opt As CloneOptions
    opt.ReplaceLocalTables = True
    opt.BatchSize = 500
    opt.Verbose = True
    opt.SkipSystemTables = True

    CloneOleDbToAccess oleDbConnString, opt
End Sub

' ---------------------------
' SCHEMA CREATION
' ---------------------------

Private Sub CreateTableWithSchema(ByVal cn As Object, ByVal tableName As String, ByRef options As CloneOptions)
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    Dim rsData As Object ' ADODB.Recordset
    Dim i As Long

    Set db = CurrentDb()

    If options.ReplaceLocalTables Then
        On Error Resume Next
        db.TableDefs.Delete tableName
        On Error GoTo 0
    Else
        If TableExists(tableName) Then Exit Sub
    End If

    ' schema-only recordset
    Set rsData = CreateObject("ADODB.Recordset")
    rsData.Open "[" & tableName & "] WHERE 1=0", cn, 1, 1  ' adOpenKeyset=1, adLockReadOnly=1

    Set tdf = db.CreateTableDef(tableName)

    For i = 0 To rsData.Fields.Count - 1
        Dim daoType As Long
        Dim txtSize As Long
        Dim adoFld As Object ' ADODB.Field

        Set adoFld = rsData.Fields(i)
        daoType = MapADOToDAOType(adoFld, options)

        Set fld = tdf.CreateField(CStr(adoFld.Name), daoType)

        If daoType = dbText Then
            If CLng(adoFld.Type) = 72 Then
                ' adGUID reports DefinedSize = 16 (its binary width), but the
                ' string form is 38 chars. Sizing from DefinedSize would
                ' silently truncate every GUID on copy.
                txtSize = GUID_TEXT_SIZE
            Else
                txtSize = CLng(Nz(adoFld.DefinedSize, 0))
                If txtSize <= 0 Or txtSize > 255 Then txtSize = 255
            End If
            fld.Size = txtSize
        End If

        If daoType = dbText Or daoType = dbMemo Then
            fld.AllowZeroLength = True
        End If

        tdf.Fields.Append fld
    Next i

    db.TableDefs.Append tdf
    db.TableDefs.Refresh

    rsData.Close
    Set rsData = Nothing

    ' Add PK where possible
    CreatePrimaryKeyForTable cn, tableName, options
End Sub

Private Sub CreatePrimaryKeyForTable(ByVal cn As Object, ByVal tableName As String, ByRef options As CloneOptions)
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    Dim rsPK As Object ' ADODB.Recordset
    Dim hasPK As Boolean

    Set db = CurrentDb()
    Set tdf = db.TableDefs(tableName)

    Set rsPK = cn.OpenSchema(28, Array(Empty, Empty, tableName)) ' adSchemaPrimaryKeys=28

    If (rsPK.EOF And rsPK.BOF) Then
        rsPK.Close
        Exit Sub
    End If

    On Error GoTo PkErr

    Set idx = tdf.CreateIndex("PK_" & tableName)
    idx.Primary = True
    idx.Unique = True

    hasPK = False
    Do While Not rsPK.EOF
        idx.Fields.Append idx.CreateField(CStr(rsPK.Fields("COLUMN_NAME").Value))
        hasPK = True
        rsPK.MoveNext
    Loop

    If hasPK Then tdf.Indexes.Append idx

    rsPK.Close
    Set rsPK = Nothing
    Exit Sub

PkErr:
    If options.Verbose Then
        Debug.Print "PK creation failed for [" & tableName & "]: " & Err.Description
    End If
    Err.Clear
    On Error GoTo 0

    On Error Resume Next
    If Not rsPK Is Nothing Then rsPK.Close
    On Error GoTo 0
End Sub

' ---------------------------
' DATA COPY
' ---------------------------

Private Sub BulkCopyTableData(ByVal cn As Object, ByVal tableName As String, ByRef options As CloneOptions)
    Dim rs As Object ' ADODB.Recordset
    Dim db As DAO.Database
    Dim rstLocal As DAO.Recordset
    Dim wrk As DAO.Workspace

    Dim i As Long
    Dim v As Variant
    Dim rowCount As Long
    Dim inTrans As Boolean
    Dim targets() As Object  ' local DAO field per source ordinal; Nothing = skip

    Set db = CurrentDb()
    Set wrk = DBEngine.Workspaces(0)

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "[" & tableName & "]", cn, 0, 1 ' adOpenForwardOnly=0, adLockReadOnly=1

    If rs.EOF And rs.BOF Then
        rs.Close
        Exit Sub
    End If

    Set rstLocal = db.OpenRecordset(tableName, dbOpenDynaset)

    ' Resolve source ordinal -> local field by NAME, once. Matching by ordinal
    ' only holds when we created the table ourselves this run; with
    ' ReplaceLocalTables = False the local table may have a different column
    ' order, which would write every value into the wrong column.
    ReDim targets(0 To rs.Fields.Count - 1)
    For i = 0 To rs.Fields.Count - 1
        On Error Resume Next
        Set targets(i) = rstLocal.Fields(CStr(rs.Fields(i).Name))
        If Err.Number <> 0 Then
            Set targets(i) = Nothing
            If options.Verbose Then
                Debug.Print "  no local column for [" & tableName & "].[" & _
                            rs.Fields(i).Name & "] -> skipped"
            End If
            Err.Clear
        End If
        On Error GoTo 0
    Next i

    On Error GoTo CopyErr

    rowCount = 0
    wrk.BeginTrans
    inTrans = True

    Do While Not rs.EOF
        rstLocal.AddNew

        For i = 0 To rs.Fields.Count - 1
            If Not targets(i) Is Nothing Then
                v = rs.Fields(i).Value
                AssignVariantToDaoField targets(i), v
            End If
        Next i

        rstLocal.Update
        rowCount = rowCount + 1

        If options.BatchSize > 0 Then
            If rowCount Mod options.BatchSize = 0 Then
                wrk.CommitTrans
                wrk.BeginTrans
            End If
        End If

        rs.MoveNext
    Loop

    wrk.CommitTrans
    inTrans = False

    rstLocal.Close
    rs.Close
    Exit Sub

CopyErr:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    ' Roll back only the uncommitted batch; earlier batches stay committed.
    If inTrans Then wrk.Rollback
    If Not rstLocal Is Nothing Then rstLocal.Close
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0

    Err.Raise errNum, "BulkCopyTableData", _
              errDesc & " (after " & rowCount & " row(s))"
End Sub

' ---------------------------
' TYPE MAPPING (ADO -> DAO)
' ---------------------------

Private Function MapADOToDAOType(ByVal f As Object, ByRef options As CloneOptions) As Long
    ' f.Type is ADODB.DataTypeEnum; numeric values are stable.
    Select Case CLng(f.Type)

        ' Integers that fit in a signed 32-bit dbLong
        ' adTinyInt=16, adSmallInt=2, adInteger=3,
        ' adUnsignedTinyInt=17, adUnsignedSmallInt=18
        Case 16, 2, 3, 17, 18
            MapADOToDAOType = dbLong

        ' Values that overflow dbLong: adUnsignedInt=19 (max 4294967295),
        ' adBigInt=20, adUnsignedBigInt=21.
        Case 19, 20, 21
            If BigIntSupported() Then
                MapADOToDAOType = DAO_dbBigInt
            Else
                ' dbDouble holds the full range and is exact up to 2^53.
                ' dbLong would silently overflow, so this is the safer fallback.
                MapADOToDAOType = dbDouble
                If options.Verbose Then
                    Debug.Print "  64-bit column [" & f.Name & "] -> dbDouble " & _
                                "(dbBigInt unavailable; values above 2^53 lose precision)"
                End If
            End If

        ' Floats / numeric
        Case 4 ' adSingle
            MapADOToDAOType = dbSingle
        Case 5, 14, 131 ' adDouble=5, adDecimal=14, adNumeric=131
            MapADOToDAOType = dbDouble

        ' Currency
        Case 6 ' adCurrency
            MapADOToDAOType = dbCurrency

        ' Boolean
        Case 11 ' adBoolean
            MapADOToDAOType = dbBoolean

        ' Dates
        Case 7, 133, 134, 135 ' adDate, adDBDate, adDBTime, adDBTimeStamp
            MapADOToDAOType = dbDate

        ' Long text
        Case 201, 203 ' adLongVarChar, adLongVarWChar
            MapADOToDAOType = dbMemo

        ' Text
        Case 129, 200, 130, 202 ' adChar, adVarChar, adWChar, adVarWChar
            MapADOToDAOType = dbText

        ' GUID
        Case 72 ' adGUID
            MapADOToDAOType = dbText

        ' Binary / image / blob
        Case 205, 128, 204 ' adLongVarBinary, adBinary, adVarBinary
            MapADOToDAOType = dbLongBinary

        Case Else
            MapADOToDAOType = dbText
    End Select
End Function

' ---------------------------
' HELPERS
' ---------------------------

Private Sub AssignVariantToDaoField(ByVal fld As DAO.Field, ByVal v As Variant)
    Dim vt As VbVarType
    Dim s As String
    Dim maxLen As Long

    If IsNull(v) Then
        fld.Value = Null
        Exit Sub
    End If

    vt = VarType(v)

    Select Case fld.Type

        Case dbText
            If vt = vbString Then
                s = CStr(v)
            ElseIf (vt And vbArray) = vbArray Then
                s = vbNullString
            Else
                s = CStr(v)
            End If

            maxLen = fld.Size
            If maxLen > 0 And Len(s) > maxLen Then s = Left$(s, maxLen)
            fld.Value = s

        Case dbMemo
            If vt = vbString Then
                fld.Value = v
            ElseIf (vt And vbArray) = vbArray Then
                fld.Value = vbNullString
            Else
                fld.Value = CStr(v)
            End If

        Case dbLongBinary
            fld.Value = v

        Case dbInteger, dbLong, dbSingle, dbDouble, dbCurrency, dbDate, dbBoolean, DAO_dbBigInt
            If vt = vbString Then
                s = Trim$(CStr(v))
                If s = vbNullString Then
                    fld.Value = Null
                    Exit Sub
                End If
            End If

            On Error GoTo ConvErr
            Select Case fld.Type
                Case dbInteger, dbLong: fld.Value = CLng(v)
                ' VBA has no native Int64; Decimal carries the full range exactly.
                Case DAO_dbBigInt: fld.Value = CDec(v)
                Case dbSingle: fld.Value = CSng(v)
                Case dbDouble: fld.Value = CDbl(v)
                Case dbCurrency: fld.Value = CCur(v)
                Case dbDate: fld.Value = CDate(v)
                Case dbBoolean: fld.Value = CBool(v)
            End Select
            On Error GoTo 0
            Exit Sub

ConvErr:
            Debug.Print "Conversion error in field [" & fld.Name & "], value='" & v & "' -> Null"
            fld.Value = Null
            Err.Clear
            On Error GoTo 0

        Case Else
            fld.Value = v
    End Select
End Sub

' dbBigInt requires a recent ACE engine AND a database file format that supports
' it, so the only reliable test is to actually create such a field. Probed once
' per session and cached.
Private Function BigIntSupported() As Boolean
    Const PROBE_TABLE As String = "~tmpBigIntProbe"

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    If mBigIntProbed Then
        BigIntSupported = mBigIntSupported
        Exit Function
    End If

    Set db = CurrentDb()

    On Error Resume Next

    db.TableDefs.Delete PROBE_TABLE   ' leftover from an aborted earlier run
    Err.Clear

    Set tdf = db.CreateTableDef(PROBE_TABLE)
    tdf.Fields.Append tdf.CreateField("v", DAO_dbBigInt)
    db.TableDefs.Append tdf

    mBigIntSupported = (Err.Number = 0)
    Err.Clear

    db.TableDefs.Delete PROBE_TABLE
    Err.Clear

    On Error GoTo 0

    mBigIntProbed = True
    BigIntSupported = mBigIntSupported
End Function

Private Function TableExists(ByVal tableName As String) As Boolean
    Dim tdf As DAO.TableDef
    For Each tdf In CurrentDb().TableDefs
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next tdf
End Function

Private Function ShouldProcessTable(ByVal tableName As String, ByRef options As CloneOptions) As Boolean
    If Len(tableName) = 0 Then Exit Function

    If options.SkipSystemTables Then
        If Left$(tableName, 4) = "MSys" Then Exit Function
        If Left$(tableName, 4) = "sys_" Then Exit Function
        If Left$(tableName, 3) = "dt_" Then Exit Function
    End If

    ShouldProcessTable = True
End Function

Private Sub ApplyDefaultOptions(ByRef options As CloneOptions)
    ' Booleans cannot be defaulted here: an unset CloneOptions is already False,
    ' so there is no way to tell "not specified" from "explicitly False".
    ' ReplaceLocalTables / SkipSystemTables / Verbose therefore default to False.    ' Booleans cannot be defaulted here: an unset CloneOptions is already False,
    ' so there is no way to tell "not specified" from "explicitly False".
    ' ReplaceLocalTables / SkipSystemTables / Verbose therefore default to False.
    ' Use CloneOleDbToAccess_Default for the opinionated clone-semantics preset.
    ' Use CloneOleDbToAccess_Default for the opinionated clone-semantics preset.
    If options.BatchSize <= 0 Then options.BatchSize = 500
End Sub
