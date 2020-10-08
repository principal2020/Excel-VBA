Sub main()
    Range("B2").AutoFilter 2, "192.168.10.1"
    With Range("A1").CurrentRegion.Offset(0, 0)
        .Resize(.Rows.Count - 1).Select
        Selection.Copy
    End With
    Sheet2.Activate
    Range("A1").Select
    ActiveSheet.Paste
    
    Sheet1.Activate
    Range("B2").AutoFilter
End Sub