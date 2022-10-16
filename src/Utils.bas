Attribute VB_Name = "Utils"
Sub checkInput()
' ���͓��e�`�F�b�N

    Sheets("�f�[�^����").Select
    
    Dim errorFlg As Boolean
    errorFlg = False
    
    For i = 3 To 14
        If Cells(i, 2).value = "�R�[�X��" Or Cells(i, 2).value = "" Or Cells(i, 3).value = "" Then
            errorFlg = True
            Exit For
        End If
    Next i
    
    If errorFlg = True Then
        Dim msgboxFlg As Integer
        msgboxFlg = MsgBox("���͂��s�����Ă��܂��B�����܂���?", vbOKCancel)
        If msgboxFlg = 2 Then
            End
        End If
    End If

End Sub

Sub resetInput()
' ���͍폜

    Sheets("�f�[�^����").Select
    
    ' ���͍폜
    Range("B3:C14").Select
    Selection.ClearContents
    
    For i = 3 To 14
        Cells(i, 2).value = "�R�[�X��"
    Next i
    
    Range("A1").Select
    
End Sub

Sub saveBook()
' �ۑ�
    
    ActiveWorkbook.Save

End Sub

Sub toDisplaySheet()
' �����L���O�y�[�W�ֈړ�
    
    Range("A1").Select
    Sheets("�����L���O").Select
    Range("A1").Select
    
End Sub

Sub clearAllData()
' �S�Ẵf�[�^�̍폜
' �f�o�b�O�p

    ' Data�V�[�g
    Sheets("Data").Select
    
    Cells(1, 9).value = "0"
    
    For rowNum = 2 To TRACK_NUM + 1
        Cells(rowNum, 2).value = "0"
        Cells(rowNum, 3).value = "0"
        Cells(rowNum, 4).value = "0"
    Next rowNum
    
    ' �����L���O�V�[�g
    deleteRanks
    
    Range("A1").Select

End Sub
