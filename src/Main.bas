Attribute VB_Name = "Main"
Sub onClickRegist()

Application.ScreenUpdating = False

' ���͓��e�s���`�F�b�N
checkInput

' �o�^�m�F�|�b�v�A�b�v
Dim putButton As Integer
putButton = MsgBox("���̓f�[�^��o�^���܂�����낵���ł��傤���H", vbYesNo)
If putButton = 7 Then
    End
End If

' ���͏��̓o�^�Ɠ��͓��e�̏���
registInput
resetInput

' �����L���O�̍X�V
setRanks
Sheets("�f�[�^����").Select

' �ۑ��m�F�|�b�v�A�b�v
putButton = MsgBox("Book��ۑ����܂����H", vbYesNo)
If putButton = 6 Then
    saveBook
End If

Application.ScreenUpdating = True

End Sub

Sub onClickReset()

Application.ScreenUpdating = False


' �폜�m�F�|�b�v�A�b�v
Dim putButton As Integer
putButton = MsgBox("���̓f�[�^���폜���܂�����낵���ł��傤���H", vbYesNo)
If putButton = 7 Then
    End
End If

resetInput

Application.ScreenUpdating = True

End Sub

Sub onClickUpdate()

Application.ScreenUpdating = False

' �����L���O�̍X�V
setRanks

' �\���V�[�g�̃Z�b�g
Sheets("�����L���O").Select

Application.ScreenUpdating = True

End Sub

Sub onClickRemoveMemo()

Application.ScreenUpdating = False

removeMEMO

Application.ScreenUpdating = True

End Sub

Sub onClickExportData()

Application.ScreenUpdating = False

exportData

Application.ScreenUpdating = True

End Sub

Sub onClickImportData()

Application.ScreenUpdating = False

importData

Application.ScreenUpdating = True
End Sub
