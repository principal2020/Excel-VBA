
  Sub replicate_and_move_files()
    Dim buf As String, cnt As Long
    Dim sDate As String
    Dim eDate As String
    Dim dt As Date

    '�V�[�g�̃p�X����n���ϐ�
    Dim dPath As String
    Dim path As String
    Dim oPath As String
    '�A�b�g�z�[���E�X�[���̕��̓V�[�g����n���ϐ�
    Dim atAnSheet As String
    Dim suAnSheet As String
    
    suAnSheet = Cells(3, 2).Value
    atAnSheet = Cells(4, 2).Value
    dPath = Cells(5, 2).Value + "\"
    path = Cells(6, 2).Value + "\"
    oPath = Cells(7, 2).Value + "\"
    '�w�肵���p�X��.xlsm�t�@�C�������ׂĎ擾
    buf = Dir(path & "*.xlsm")
    '�O��t�@�C�����ő�100���Ή�
    Dim fileName(100) As String
    Dim fileReName(100) As String
    Dim seDate As String
    Dim fileDate As String
    Dim rn As Integer
    ReDim intRains(5) As Integer
    
    '-----Suumo,Athome���̓V�[�g����l���擾�B-----
    Workbooks.Open dPath + suAnSheet
    ReDim suStr(100) As String
    ReDim suPrice(100) As Integer
    ReDim suEng(100) As Integer
    Dim rCnt As Integer
    
    rCnt = 0
    suPrice(0) = 9999
    suStr(0) = "substring"
     
     '�z��ɃX�[�����̓V�[�g�̏��ݒn�ŏ���5�����A���i���i�[����B
    Do While Trim(suStr(rCnt)) <> ""
        If suStr(0) <> "substring" Then
            rCnt = rCnt + 1
        End If
        suStr(rCnt) = Left(ActiveSheet.Cells(8 + rCnt, 17).Value, 5)
        
        If InStr(suStr(rCnt), "�厚") <> 0 Then
            suStr(rCnt) = Replace(suStr(rCnt), "�厚", "") + Mid(ActiveSheet.Cells(8 + rCnt, 17).Value, 6, 2)
        End If

        If ActiveSheet.Cells(8 + rCnt, 34).Value <> "" Then
            suPrice(rCnt) = Int(Replace(ActiveSheet.Cells(8 + rCnt, 34).Value, "���~", ""))
        End If
        
        suEng(rCnt) = ActiveSheet.Cells(8 + rCnt, 66).Value
        
        If rCnt > 100 Then
            Exit Do
        End If
    Loop
    
    ReDim Preserve suStr(rCnt - 1)
    ReDim Preserve suPrice(rCnt - 1)
    ReDim Preserve suEng(rCnt - 1)
    'MsgBox (engNum)
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    'Exit Sub
    
    Workbooks.Open dPath + atAnSheet
    ReDim atStr(200) As String
    ReDim atPrice(200) As Integer
    ReDim ateng(200) As Integer
    'Dim engNum As Integer
    'Dim rCnt As Integer
    rCnt = 0
    'Do While ActiveSheet.Cells(8 + rCnt, 17).Value <> ""
    atPrice(0) = 9999
    atStr(0) = "atbstring"
     
     '�z��ɃA�b�g�z�[�����̓V�[�g�̏��ݒn�ŏ���5�����A���i���i�[����B
     Do While Trim(atStr(rCnt)) <> ""
        If atStr(0) <> "atbstring" Then
            rCnt = rCnt + 1
        End If
        atStr(rCnt) = Mid(ActiveSheet.Cells(2 + rCnt, 5).Value, 4, 5)
        If InStr(atStr(rCnt), "�厚") <> 0 Then
                atStr(rCnt) = Replace(atStr(rCnt), "�厚", "") + Mid(ActiveSheet.Cells(2 + rCnt, 5).Value, 9, 2)
            Else
            'no process
        End If

        If ActiveSheet.Cells(2 + rCnt, 7).Value <> "" Then
            atPrice(rCnt) = ActiveSheet.Cells(2 + rCnt, 7).Value
        End If
        
        ateng(rCnt) = ActiveSheet.Cells(2 + rCnt, 14).Value
        '�������[�v���p
        If rCnt > 100 Then
            Exit Do
        End If
    Loop
    ReDim Preserve atStr(rCnt - 1)
    ReDim Preserve atPrice(rCnt - 1)
    ReDim Preserve ateng(rCnt - 1)
    ActiveWorkbook.Save
    ActiveWorkbook.Close
    
    '�t�@�C��������
    Do While buf <> ""
        cnt = cnt + 1
        fileName(cnt) = buf
        If Len(fileName(cnt)) >= 14 Then
            
            Workbooks.Open path + fileName(cnt)
            
            '�X���b�V�����܂ޓ��t��string�^��yyyymmdd�ɕϊ�
            With ActiveSheet
                sDate = .Cells(37, 31).Value
                sDate = DateAdd("d", 1, sDate)
                sDate = Format(sDate, "mm��dd��")
                
                'dt��13�����Z
                eDate = DateAdd("d", 13, sDate)
                eDate = Format(eDate, "mm��dd��")
                'MsgBox (eDate)
                '�J�n���A�I������A��(�t�@�C�����p�ϐ�)
                seDate = sDate + "�`" + eDate
            End With
            
            '�t�@�C�����J�����ƂɃ��C���Y�p�̗����z���������
            For i = 0 To 2
                    intRains(i) = 0
            Next
            '���C���Y�̌��݂̌�����-3�`3�̗��������Z
            If Cells(31, 18).Value <> "" Then
                rn = randInt()
                intRains(0) = Cells(31, 18).Value + rn
            End If
            If Cells(31, 23).Value <> "" Then
                rn = randInt()
                intRains(1) = Cells(31, 23).Value + rn
            End If
            If Cells(31, 28).Value <> "" Then
                rn = randInt()
                intRains(2) = Cells(31, 28).Value + rn
            End If
            
            '�R�s�[���t�@�C��������yyyymmdd�`mmdd.xlsm�̕������폜
            fileReName(cnt) = _
            Left(fileName(cnt), Len(fileName(cnt)) - 18)
            
            '�R�s�[��t�@�C�������R�s�[���t�@�C���� + mmdd�`mmdd.xlsm�ɐݒ�
            '�}�N���L���u�b�N���ʏ�̃u�b�N�ɂ��邩�͂܂����߂Ă��Ȃ�
            fileReName(cnt) = _
            fileReName(cnt) + seDate + ".xlsm"

            ActiveWorkbook.Save
            ActiveWorkbook.Close
            
            FileCopy path + fileName(cnt), oPath + fileReName(cnt)
            Workbooks.Open oPath + fileReName(cnt)
            
            Dim lCnt As Integer
            Dim tempProName As String
            ReDim tempPrice(5) As Integer
            lCnt = 0
            
            If InStr(tempProName, "�厚") = 0 Then
                tempProName = Replace(Cells(23, 9).Value, "�厚", "")
            Else
                tempProName = Cells(23, 9).Value
            End If
            
            tempProName = Left(tempProName, 5)
            
            If ActiveSheet.Cells(24, 9).Value <> "" Then
                tempPrice(0) = Replace(ActiveSheet.Cells(24, 9).Value, "���~", "")
            End If
            
            If ActiveSheet.Cells(24, 15).Value <> "" Then
                tempPrice(1) = Replace(ActiveSheet.Cells(24, 15).Value, "���~", "")
                lCnt = lCnt + 1
            End If
            
            If ActiveSheet.Cells(24, 20).Value <> "" Then
                tempPrice(2) = Replace(ActiveSheet.Cells(24, 20).Value, "���~", "")
                lCnt = lCnt + 1
            End If
            ReDim Preserve tempPrice(5)
            
            '----�t�@�C���I�[�v������while���Ŏ��s--------------------------------
            Dim tstProName As String
            Dim sCnt As Integer
            Dim smCnt As Integer
            ReDim sudeteng(5) As Integer
            sCnt = 0
            smCnt = 0
            For i = 0 To lCnt
                sCnt = 0
                For Each suCompStr In suStr()
                    If suCompStr = tempProName And suPrice(sCnt) = tempPrice(i) Then
                        '�e�X�g�p
                        'MsgBox ("�����ƃ}�b�`����Z������������܂����B")
                       sudeteng(i) = suEng(sCnt)
                       smCnt = smCnt + 1
                        If i = 0 Then
                            sCnt = sCnt + 1
                            GoTo continue
                        
                        ElseIf i = 1 And smCnt = 3 And tempPrice(i - 1) = tempPrice(i) Then
                            sCnt = sCnt + 1
                            GoTo continue
                        ElseIf i = 2 And smCnt = 6 And tempPrice(i - 1) = tempPrice(i) Then
                            sCnt = sCnt + 1
                            GoTo continue
                        Else
                            'no process
                        End If
                    End If
                    sCnt = sCnt + 1
                Next
continue:
            Next
            ReDim Preserve sudeteng(2)
            
            Dim aCnt As Integer
            Dim amCnt As Integer
            ReDim atdeteng(5) As Integer
            'atDetEng��Empty�l�ɂȂ�Ȃ��悤�ɂ���B
            For i = 0 To 2
                atdeteng(i) = 0
            Next i
            aCnt = 0
            amCnt = 0
            For i = 0 To lCnt
                aCnt = 0
                For Each atCompStr In atStr()
                    If atCompStr = tempProName And atPrice(aCnt) = tempPrice(i) Then
                        '�e�X�g�p
                        'MsgBox ("athome name�����ƃ}�b�`����Z������������܂���")
                        atdeteng(i) = ateng(aCnt)
                        amCnt = amCnt + 1
                        
                        If i = 0 Then
                            aCnt = aCnt + 1
                            GoTo continuea
                        
                        ElseIf i = 1 And amCnt = 3 And tempPrice(i - 1) = tempPrice(i) Then
                            aCnt = aCnt + 1
                            GoTo continuea
                        ElseIf i = 2 And amCnt = 6 And tempPrice(i - 1) = tempPrice(i) Then
                            aCnt = aCnt + 1
                            GoTo continuea
                        Else
                            'no process
                        End If
                    End If
                    aCnt = aCnt + 1
                Next
continuea:
            Next i
            ReDim Preserve atdeteng(2)
            
            '-------------------------------------------------------------------
            If atdeteng(0) > 0 Then
                Select Case lCnt
                    Case 0
                        If atdeteng(0) > 0 Then
                            ActiveSheet.Cells(33, 18).Value = atdeteng(0)
                        End If
                    Case 1
                        If atdeteng(0) > 0 Then
                            ActiveSheet.Cells(33, 18).Value = atdeteng(0)
                        End If
                        If atdeteng(1) > 0 Then
                            ActiveSheet.Cells(33, 23).Value = atdeteng(1)
                        End If
                    Case 2
                        If atdeteng(0) > 0 Then
                            ActiveSheet.Cells(33, 18).Value = atdeteng(0)
                        End If
                        If atdeteng(1) > 0 Then
                            ActiveSheet.Cells(33, 23).Value = atdeteng(1)
                        End If
                        If atdeteng(2) > 0 Then
                            ActiveSheet.Cells(33, 28).Value = atdeteng(2)
                        End If
                    'Case 3
                        'ActiveSheet.Cells(33, 18).Value = atDetEng(0)
                End Select
            End If
            If sudeteng(0) > 0 Then
                Select Case lCnt
                    Case 0
                        If sudeteng(0) > 0 Then
                            ActiveSheet.Cells(35, 18).Value = sudeteng(0)
                        End If
                    Case 1
                        If sudeteng(0) > 0 Then
                            ActiveSheet.Cells(35, 18).Value = sudeteng(0)
                        End If
                        If sudeteng(1) > 0 Then
                            ActiveSheet.Cells(35, 23).Value = sudeteng(1)
                        End If
                    Case 2
                        If sudeteng(0) > 0 Then
                            ActiveSheet.Cells(35, 18).Value = sudeteng(0)
                        End If
                        If sudeteng(1) > 0 Then
                            ActiveSheet.Cells(35, 23).Value = sudeteng(1)
                        End If
                        If sudeteng(2) > 0 Then
                            ActiveSheet.Cells(35, 28).Value = sudeteng(2)
                        End If
                End Select
            End If
            
            ActiveSheet.Cells(34, 31).Value = Format(sDate, "yyyy/mm/dd")
            ActiveSheet.Cells(37, 31).Value = Format(eDate, "yyyy/mm/dd")
            ActiveSheet.Cells(1, 2).Value = Date
            If intRains(0) <> 0 And IsEmpty(intRains(0)) = False Then
                ActiveSheet.Cells(31, 18).Value = intRains(0)
            End If
            If intRains(1) <> 0 And IsEmpty(intRains(1)) = False Then
                ActiveSheet.Cells(31, 23).Value = intRains(1)
            End If
            If intRains(2) <> 0 And IsEmpty(intRains(2)) = False Then
                ActiveSheet.Cells(31, 28).Value = intRains(2)
            End If
            ActiveWorkbook.Save
            ActiveWorkbook.Close
            buf = Dir()
        '�t�@�C�����̕���������18������菭�Ȃ��ꍇ�̏���
        Else
            MsgBox ("�t�@�C�������s���ł��B")
            buf = Dir()
        End If
    Loop
     
End Sub
'�t�@�C���̈��
Sub printSheet()
    Dim rc As Integer
    Dim buf As String
    Dim path As String
    Dim fileName As String
    path = Cells(7, 2).Value + "\"
    rc = MsgBox("�V�[�g��������܂����H", vbYesNo + vbQuestion, "�m�F")
    If rc = vbYes Then
        buf = Dir(path & "*.xlsm")
        Do While buf <> ""
            If buf = "" Then
                Exit Sub
            End If
            fileName = buf
            Workbooks.Open path + fileName
            'MsgBox (path + fileName)
            
            ActiveWorkbook.ActiveSheet.PrintOut
            ActiveWorkbook.Save
            ActiveWorkbook.Close
            buf = Dir()
        Loop
    Else
        'no process
    End If
End Sub

Function randInt() As Integer
    '0�ȊO�̗��������������܂Ń��[�v
    Dim rn As Integer
    rn = 0
    Do While rn = 0
        Randomize
        rn = Int((3 - (-3) + 1) * Rnd + (-3))
    Loop
    randInt = rn
End Function
