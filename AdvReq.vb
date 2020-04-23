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
    thisSheet = "�L���f�ڏ��������̓t�H�[��"
    singSheet = "�L���f�ڏ�����(�P��)"
    multiSheet = "�L���f�ڏ�����(����)"
    proType = ComboType.Value
    If proType = "" Then
        errString = "�������"
    End If
    comName = setCellsValues("�ƎҖ�", 5, 4, errString)
    oRepName = setCellsValues("�S����", 7, 4, errString)
    proName1 = setCellsValues("�������P�@������", 10, 4, errString)
    proPrice1 = setCellsValues("�����P�@���i", 12, 4, errString)
    If Trim(Cells(14, 4).Value) <> "" Or Trim(Cells(16, 4).Value) <> "" Then
        proName2 = setCellsValues("�����Q�@������", 14, 4, errString)
        proPrice2 = setCellsValues("�����Q�@���i", 16, 4, errString)
    End If
    If Trim(Cells(14, 4).Value) <> "" And Trim(Cells(16, 4).Value) <> "" Then
        If Trim(Cells(18, 4).Value) <> "" Or Trim(Cells(20, 4).Value) <> "" Then
            proName3 = setCellsValues("�����R�@������", 18, 4, errString)
            proPrice3 = setCellsValues("�����R�@���i", 20, 4, errString)
        End If
    End If
    remarks = setCellsValues("���l", 22, 4, errString)
    iRepName = "�S���ҁi�@" + setCellsValues("���ВS����", 24, 4, errString)
    If errString <> "" Then
        MsgBox (errString & vbCrLf & "�͕K�{���ڂł��B")
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
            .Cells(42, 8).Value = iRepName + "�@�j"
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
            .Cells(38, 9).Value = iRepName + "�@�j"
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
        inpName = InputBox("�S���Җ�����͂��Ă��������B�s���̏ꍇ��OK�������Ă��������B")
    Else
        inpName = ioRepName
    End If
    With Worksheets(s)
        .Cells(i, b).Value = inpName
        If Trim(.Cells(i, b).Value) = "�l" Then
            .Cells(i, a).Value = "���S����:"
            .Cells(i, b).Value = ""
            .Cells(i, c).Value = "�l"
            Range(ranA).Borders(xlEdgeBottom).LineStyle = False
            Range(ranB).Borders(xlEdgeBottom).LineStyle = xlContinuous
            'MsgBox "�S���Җ�����͂��Ă��������B�s���̏ꍇ�͂�����x�{�^���������Ă��������B"
            .Cells(i, b).Select
            Exit Function
        ElseIf Trim(.Cells(i, b).Value) = "" Then
            .Cells(i, a).Value = "���S����"
            .Cells(i, b).Value = "�l"
            .Cells(i, c).Value = ""
            .Range(ranA).Borders(xlEdgeBottom).LineStyle = False
            .Range(ranB).Borders(xlEdgeBottom).LineStyle = False
        ElseIf Trim(.Cells(i, b).Value) <> "�l" And Trim(.Cells(i, b).Value) <> "" Then
            .Cells(i, a).Value = "���S����:"
            .Cells(i, c).Value = "�l"
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
            Case "���Ìˌ�"
                proType = ChrW(&H2611) + "���Ìˌ��@�@�@ ���V�z�ˌ�" & vbCrLf & "�����Ã}���V���� ���y�n"
            Case "�V�z�ˌ�"
                proType = "�����Ìˌ��@�@�@ " + ChrW(&H2611) + "�V�z�ˌ�" & vbCrLf & "�����Ã}���V���� ���y�n"
            Case "���Ã}���V����"
                proType = "�����Ìˌ��@�@�@ ���V�z�ˌ�" & vbCrLf & ChrW(&H2611) + "���Ã}���V���� ���y�n"
            Case "�y�n"
                proType = "�����Ìˌ��@�@�@ ���V�z�ˌ�" & vbCrLf & "�����Ã}���V���� " + ChrW(&H2611) + "�y�n"
            Case Else
                MsgBox ("������ʂ����X�g����I�����Ă��������B")
        End Select
    Else
        Select Case comVal
            Case "���Ìˌ�"
                proType = ChrW(&H2611) + "���Ìˌ��@�@���V�z�ˌ��@�@�����Ã}���V�����@�@���y�n"
            Case "�V�z�ˌ�"
                proType = "�����Ìˌ��@�@" + ChrW(&H2611) + "�V�z�ˌ��@�@�����Ã}���V�����@�@���y�n"
            Case "���Ã}���V����"
                proType = "�����Ìˌ��@�@���V�z�ˌ��@�@" + ChrW(&H2611) + "���Ã}���V�����@�@���y�n"
            Case "�y�n"
                proType = "�����Ìˌ��@�@���V�z�ˌ��@�@�����Ã}���V�����@�@" + ChrW(&H2611) + "�y�n"
            Case Else
                MsgBox ("������ʂ����X�g����I�����Ă��������B")
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
        'MsgBox (cellsName + "�͓��͕K�{�t�B�[���h�ł��B")
        If errString = "" And cellsName <> "���l" And cellsName <> "�S����" Then
            errString = cellsName
        ElseIf errString <> "" And cellsName <> "���l" And cellsName <> "�S����" Then
            errString = errString + "�C" + cellsName
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
    rc = MsgBox("�V�[�g��������܂����H", vbYesNo + vbQuestion, "�m�F")
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
          ComboType.AddItem ("���Ìˌ�")
          ComboType.AddItem ("�V�z�ˌ�")
          ComboType.AddItem ("���Ã}���V����")
          ComboType.AddItem ("�y�n")
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
