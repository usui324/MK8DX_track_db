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

Sub inputSampleData()
' �T���v���f�[�^���Z�b�g����
' �f�o�b�O�p
    Dim rootPath, filePath
    Dim str As String

    ' �T���v���f�[�^�t�@�C���p�X
    rootPath = ThisWorkbook.Path
    filePath = rootPath & "\sampleData\sampleData.txt"
    
    ' �f�[�^��ǂݍ���
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        str = .ReadText
        .Close
    End With
    
    ' �N���b�v�{�[�h�ɒl���Z�b�g
    Dim dataObj As DataObject
    Set dataObj = New DataObject
    dataObj.SetText str
    dataObj.PutInClipboard
    Set dataObj = Nothing
    
    ' Data�V�[�g�ɓ\��t��
    ' NOTE: �J���}��؂�̃f�[�^���󂯕t����悤�ɐݒ肷��K�v������
    Sheets("Data").Select
    Range("B2").Select
    ActiveSheet.Paste
    
    Range("A1").Select
    
End Sub

Sub exportData()
' �f�[�^��txt�t�@�C���ɃG�N�X�|�[�g
    ' �G�N�X�|�[�g�t�@�C�����w��
    ChDir ThisWorkbook.Path
    Dim saveFileName As String
    saveFileName = Application.GetSaveAsFilename(InitialFileName:="trackData.txt", filefilter:="�R�[�X�f�[�^,*.txt")

    ' �L�����Z������
    If saveFileName = "False" Then
        Exit Sub
    End If

    ' �o�͂���ΏۃV�[�g
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")
    
    ' �t�@�C���ɏ�������
    Open saveFileName For Output As #1
    ' �͋[��
    Print #1, ws.Cells(1, 8).value & "," & ws.Cells(1, 9).value
    ' trackData
    Dim i As Integer
    For i = 1 To TRACK_NUM
        Print #1, printLine(ws, i + 1)
    Next i
    
    Close #1
    
    MsgBox saveFileName & "�Ƀf�[�^���o�͂��܂���", vbInformation

End Sub

Function printLine(ws As Worksheet, rowNo As Integer) As String
' �t�@�C���o�͂����s�̕������Ԃ�
    Dim trackName As String
    Dim rankSum As String
    Dim pointSum As String
    Dim raceNum As String
    
    trackName = ws.Cells(rowNo, 1).value
    rankSum = ws.Cells(rowNo, 2).value
    pointSum = ws.Cells(rowNo, 3).value
    raceNum = ws.Cells(rowNo, 4).value
    
    Dim str As String
    printLine = trackName & "," & rankSum & "," & pointSum & "," & raceNum
    
End Function

Sub importData()
' txt�t�@�C������f�[�^���C���|�[�g
    Dim openFileName As String
    Dim ws As Worksheet
    Dim line As String
    Dim arr As Variant
    Dim rowNo As Integer
    Dim columnNo As Integer

    ' �C���|�[�g�t�@�C�����w��
    ChDir ThisWorkbook.Path
    openFileName = Application.GetOpenFilename("�R�[�X�f�[�^,*.txt", , "�C���|�[�g����f�[�^�t�@�C�����w��")
    
    ' �L�����Z������
    If openFileName = "False" Then
        Exit Sub
    End If
    
    ' ���͑ΏۃV�[�g
    Set ws = ThisWorkbook.Worksheets("Data")
    
    ' �͋[�񐔂̓���
    Open openFileName For Input As #1
    Line Input #1, line
    arr = Split(line, ",")
    ws.Cells(1, 9).value = arr(1)
    
    ' �R�[�X�f�[�^�̓���
    rowNo = 2
    While Not EOF(1)
        Line Input #1, line
        arr = Split(line, ",")
        
        ' �z�񒷂�5�ȏ�̏ꍇ�̓G���[
        If UBound(arr) >= 4 Then
            MsgBox "�f�[�^���s���ł�", vbExclamation
            Exit Sub
        End If
        
        For columnNo = LBound(arr) To UBound(arr)
            ws.Cells(rowNo, columnNo + 1).value = arr(columnNo)
        Next columnNo
        rowNo = rowNo + 1
    Wend
    Close #1
    
    ' �����L���O�f�[�^���X�V
    setRanks
    
    MsgBox "�f�[�^���C���|�[�g���܂���", vbInformation
End Sub
