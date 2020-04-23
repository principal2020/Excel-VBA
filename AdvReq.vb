Private Sub CommandButton1_Click()
  Call clearInputValues
End Sub

Private Sub Done_Click()
    Dim comName As String
    Dim oRepName As String
    Dim proName1 As String
    Dim proName2 As String
    Dim proName3 As String
    Dim proPrice1 As String
    Dim proPrice2 As String
    Dim proPrice3 As String
    Dim remarks As String
    Dim iRepName As String
    Dim thisSheet As String
    Dim singSheet As String
    Dim multiSheet As String
    Dim errString As String
    
    errString = ""
    thisSheet = "広告掲載承諾書入力フォーム"
    singSheet = "広告掲載承諾書(単数)"
    multiSheet = "広告掲載承諾書(複数)"
    proType = ComboType.Value
    If proType = "" Then
        errString = "物件種別"
    End If
    comName = setCellsValues("業者名", 5, 4, errString)
    oRepName = setCellsValues("担当者", 7, 4, errString)
    proName1 = setCellsValues("物件名１　物件名", 10, 4, errString)
    proPrice1 = setCellsValues("物件１　価格", 12, 4, errString)
    If Trim(Cells(14, 4).Value) <> "" Or Trim(Cells(16, 4).Value) <> "" Then
        proName2 = setCellsValues("物件２　物件名", 14, 4, errString)
        proPrice2 = setCellsValues("物件２　価格", 16, 4, errString)
    End If
    If Trim(Cells(14, 4).Value) <> "" And Trim(Cells(16, 4).Value) <> "" Then
        If Trim(Cells(18, 4).Value) <> "" Or Trim(Cells(20, 4).Value) <> "" Then
            proName3 = setCellsValues("物件３　物件名", 18, 4, errString)
            proPrice3 = setCellsValues("物件３　価格", 20, 4, errString)
        End If
    End If
    remarks = setCellsValues("備考", 22, 4, errString)
    iRepName = "担当者（　" + setCellsValues("弊社担当者", 24, 4, errString)
    If errString <> "" Then
        MsgBox (errString & vbCrLf & "は必須項目です。")
        Exit Sub
    End If
    If proName2 <> "" And IsNull(proPrice2) = False Or proName3 <> "" And IsNull(proPrice3) = False Then
        With Worksheets(multiSheet)
            proType = setProType(2)
            'Call chgFontSize(comName, multiSheet)
            .Cells(8, 2).Value = comName
            Call chgFontSize(8, 2, multiSheet)
            Call repName(11, 2, 3, 4, oRepName, multiSheet, "B11", "C11")
            .Cells(19, 4).Value = proType
            .Cells(21, 4).Value = proName1
            .Cells(23, 4).Value = proName2
            .Cells(25, 4).Value = proName3
            .Cells(21, 9).Value = proPrice1
            .Cells(23, 9).Value = proPrice2
            .Cells(25, 9).Value = proPrice3
            .Cells(27, 4).Value = remarks
            .Cells(42, 8).Value = iRepName + "　）"
            Call printSheet(multiSheet)
        End With
    Else
        proType = setProType(1)
        Worksheets(singSheet).Activate
        With Worksheets(singSheet)
            'Call chgFontSize(comName, singSheet)
            .Cells(8, 2).Value = comName
            Call chgFontSize(8, 2, singSheet)
            Call repName(11, 2, 3, 4, oRepName, singSheet, "B11", "C11")
            .Cells(19, 4).Value = proType
            .Cells(21, 4).Value = proName1
            .Cells(19, 8).Value = proPrice1
            .Cells(23, 4).Value = remarks
            .Cells(38, 9).Value = iRepName + "　）"
            Call printSheet(singSheet)
        End With
    End If
    ActiveSheet.Activate
    
    Worksheets(thisSheet).Activate
    Worksheets(thisSheet).Cells(5, 4).Select
    ActiveWorkbook.Save
End Sub

'Description: Set representive persons name.
Function repName(ByRef rowNo, colRepA, colRepB, colRepC, ioRepName As String, sheetName As String, celRangA, celRangB)
    Dim i As Integer
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim s As String
    Dim ranA As String
    Dim ranB As String
    Dim inpName As String
    i = rowNo
    a = colRepA
    b = colRepB
    c = colRepC
    s = sheetName
    ranA = celRangA
    ranB = celRangB
    If ioRepName = "" Then
        inpName = InputBox("担当者名を入力してください。不明の場合はOKを押してください。")
    Else
        inpName = ioRepName
    End If
    With Worksheets(s)
        .Cells(i, b).Value = inpName
        If Trim(.Cells(i, b).Value) = "様" Then
            .Cells(i, a).Value = "ご担当者:"
            .Cells(i, b).Value = ""
            .Cells(i, c).Value = "様"
            Range(ranA).Borders(xlEdgeBottom).LineStyle = False
            Range(ranB).Borders(xlEdgeBottom).LineStyle = xlContinuous
            'MsgBox "担当者名を入力してください。不明の場合はもう一度ボタンを押してください。"
            .Cells(i, b).Select
            Exit Function
        ElseIf Trim(.Cells(i, b).Value) = "" Then
            .Cells(i, a).Value = "ご担当者"
            .Cells(i, b).Value = "様"
            .Cells(i, c).Value = ""
            .Range(ranA).Borders(xlEdgeBottom).LineStyle = False
            .Range(ranB).Borders(xlEdgeBottom).LineStyle = False
        ElseIf Trim(.Cells(i, b).Value) <> "様" And Trim(.Cells(i, b).Value) <> "" Then
            .Cells(i, a).Value = "ご担当者:"
            .Cells(i, c).Value = "様"
            .Range(ranA).Borders(xlEdgeBottom).LineStyle = False
            .Range(ranB).Borders(xlEdgeBottom).LineStyle = xlContinuous
        Else
            'no process
        End If
    End With
End Function

'Select one of the value from listbox.
Function setProType(intJudge As Integer) As String
    Dim comVal As String
    Dim proType As String
    comVal = ComboType.Value
    If intJudge = 1 Then
        Select Case comVal
            Case "中古戸建"
                proType = ChrW(&H2611) + "中古戸建　　　 □新築戸建" & vbCrLf & "□中古マンション □土地"
            Case "新築戸建"
                proType = "□中古戸建　　　 " + ChrW(&H2611) + "新築戸建" & vbCrLf & "□中古マンション □土地"
            Case "中古マンション"
                proType = "□中古戸建　　　 □新築戸建" & vbCrLf & ChrW(&H2611) + "中古マンション □土地"
            Case "土地"
                proType = "□中古戸建　　　 □新築戸建" & vbCrLf & "□中古マンション " + ChrW(&H2611) + "土地"
            Case Else
                MsgBox ("物件種別をリストから選択してください。")
        End Select
    Else
        Select Case comVal
            Case "中古戸建"
                proType = ChrW(&H2611) + "中古戸建　　□新築戸建　　□中古マンション　　□土地"
            Case "新築戸建"
                proType = "□中古戸建　　" + ChrW(&H2611) + "新築戸建　　□中古マンション　　□土地"
            Case "中古マンション"
                proType = "□中古戸建　　□新築戸建　　" + ChrW(&H2611) + "中古マンション　　□土地"
            Case "土地"
                proType = "□中古戸建　　□新築戸建　　□中古マンション　　" + ChrW(&H2611) + "土地"
            Case Else
                MsgBox ("物件種別をリストから選択してください。")
        End Select
    End If
    setProType = proType
End Function

'Set cells values to varriants(Notice:String type only).
Function setCellsValues(ByRef cellsName As String, inpRow As Integer, inpCol As Integer, errString) As String
    Dim cellsValue As String
    cellsValue = ActiveSheet.Cells(inpRow, inpCol).Value
    If IsNumeric(cellsValue) = True Then
        cellsValue = Str(cellsValue)
    End If
    If cellsValue <> "" Then
        setCellsValues = cellsValue
    ElseIf cellsValue = "" Then
        'MsgBox (cellsName + "は入力必須フィールドです。")
        If errString = "" And cellsName <> "備考" And cellsName <> "担当者" Then
            errString = cellsName
        ElseIf errString <> "" And cellsName <> "備考" And cellsName <> "担当者" Then
            errString = errString + "，" + cellsName
            'MsgBox (errString)
            ActiveSheet.Cells(inpRow, inpCol).Select
        Else
            'no process
        End If
    End If
End Function

Function setCellsPrice(cellsName As String, inpRow As Integer, inpCol As Integer) As Integer
    Dim cellsValue As Integer
    cellsValue = ActiveSheet.Cells(inpRow, inpCol).Value
    If IsNull(cellsValue) = False Then
        setCellsPrice = cellsValue
    ElseIf cellsName = "" Then
        errString = errString + cellsName
        MsgBox (errString)
        ActiveSheet.Cells(inpRow, inpCol).Select
    End If
End Function

Function printSheet(sheetName As String)
    Dim rc As Integer
    rc = MsgBox("シートを印刷しますか？", vbYesNo + vbQuestion, "確認")
    If rc = vbYes Then
        Worksheets(sheetName).PrintOut
    Else
        'no process
    End If
End Function

Function clearInputValues()
    Cells(5, 4).Value = ""
    Cells(7, 4).Value = ""
    ComboType.Value = ""
    Cells(10, 4).Value = ""
    Cells(12, 4).Value = ""
    Cells(14, 4).Value = ""
    Cells(16, 4).Value = ""
    Cells(18, 4).Value = ""
    Cells(20, 4).Value = ""
    Cells(22, 4).Value = ""
    Cells(24, 4).Value = ""
End Function

Function setProtypeValue()
    If ComboType.ListCount < 1 Then
          ComboType.Clear
          ComboType.AddItem ("")
          ComboType.AddItem ("中古戸建")
          ComboType.AddItem ("新築戸建")
          ComboType.AddItem ("中古マンション")
          ComboType.AddItem ("土地")
          ComboType.Style = fmStyleDropDownList
    End If
End Function


'Define contants of combobox
Private Sub ComboType_DropButtonClick()
    Call setProtypeValue
End Sub

Private Sub ComboType_GotFocus()
    Call setProtypeValue
End Sub


Private Sub chgFontSize(comRow As Integer, comCol As Integer, sheetName As String)
    With Worksheets(sheetName)
        If Len(.Cells(comRow, comCol).Value) >= 20 Then
            .Cells(comRow, comCol).Font.Size = 14
        ElseIf Len(.Cells(comRow, comCol).Value) <= 10 Then
            .Cells(comRow, comCol).Font.Size = 20
        Else
            .Cells(comRow, comCol).Font.Size = 16
        End If
    End With
End Sub
