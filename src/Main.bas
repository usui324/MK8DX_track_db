Attribute VB_Name = "Main"
Sub onClickRegist()

Application.ScreenUpdating = False

' 入力内容不備チェック
checkInput

' 登録確認ポップアップ
Dim putButton As Integer
putButton = MsgBox("入力データを登録しますがよろしいでしょうか？", vbYesNo)
If putButton = 7 Then
    End
End If

' 入力情報の登録と入力内容の消去
registInput
resetInput

' ランキングの更新
setRanks
Sheets("データ入力").Select

' 保存確認ポップアップ
putButton = MsgBox("Bookを保存しますか？", vbYesNo)
If putButton = 6 Then
    saveBook
End If

Application.ScreenUpdating = True

End Sub

Sub onClickReset()

Application.ScreenUpdating = False


' 削除確認ポップアップ
Dim putButton As Integer
putButton = MsgBox("入力データを削除しますがよろしいでしょうか？", vbYesNo)
If putButton = 7 Then
    End
End If

resetInput

Application.ScreenUpdating = True

End Sub

Sub onClickUpdate()

Application.ScreenUpdating = False

' ランキングの更新
setRanks

' 表示シートのセット
Sheets("ランキング").Select

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
