Public Sub RunClone_Custom()
    Dim conn As String
    conn = "Provider=...;Data Source=...;User ID=...;Password=...;"

    Dim opt As CloneOptions
    opt.ReplaceLocalTables = True
    opt.BatchSize = 500
    opt.Verbose = True
    opt.SkipSystemTables = True

    CloneOleDbToAccess conn, opt
End Sub
