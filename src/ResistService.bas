Attribute VB_Name = "ResistService"
' �f�[�^�o�^
Sub registInput()

' �e�R�[�X�ɑ΂��ē������\�b�h�����s
For i = 3 To 14

    ' �R�[�X���E���ʖ����͗�̓X�L�b�v
    Sheets("�f�[�^����").Select
    
    If Cells(i, 2).value = "�R�[�X��" Or Cells(i, 3).value = "" Then
        GoTo CONTINUE1:
        
    End If
    
    Dim trackName As String
    trackName = Cells(i, 2).value
    Dim rankValue As Integer
    rankValue = Cells(i, 3).value
    
    ' ���ʂɑ΂��链�_���v�Z
    Sheets("MasterData").Select
    Dim pointValue As Integer
    pointValue = Cells(rankValue + 1, 4).value

    ' Data����R�[�X���ɍ��v������T��
    Dim trackRow As Integer
    
    Sheets("Data").Select
    For j = 2 To 97
        If Cells(j, 1).value = trackName Then
            trackRow = j
            GoTo CONTINUE2:
        End If
CONTINUE2:
    Next j
    
    ' �f�[�^�̏�������
    Sheets("Data").Select
    Cells(trackRow, 2).value = Cells(trackRow, 2).value + rankValue
    Cells(trackRow, 3).value = Cells(trackRow, 3).value + pointValue
    Cells(trackRow, 4).value = Cells(trackRow, 4).value + 1
    
CONTINUE1:
Next i

' �͋[�񐔂̉��Z
Sheets("Data").Select
Cells(1, 9).value = Cells(1, 9).value + 1

Sheets("�f�[�^����").Select
Range("A1").Select

End Sub

