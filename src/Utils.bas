Attribute VB_Name = "Utils"
Sub checkInput()
' 入力内容チェック

    Sheets("データ入力").Select
    
    Dim errorFlg As Boolean
    errorFlg = False
    
    For i = 3 To 14
        If Cells(i, 2).value = "コース名" Or Cells(i, 2).value = "" Or Cells(i, 3).value = "" Then
            errorFlg = True
            Exit For
        End If
    Next i
    
    If errorFlg = True Then
        Dim msgboxFlg As Integer
        msgboxFlg = MsgBox("入力が不足しています。続けますか?", vbOKCancel)
        If msgboxFlg = 2 Then
            End
        End If
    End If

End Sub

Sub resetInput()
' 入力削除

    Sheets("データ入力").Select
    
    ' 入力削除
    Range("B3:C14").Select
    Selection.ClearContents
    
    For i = 3 To 14
        Cells(i, 2).value = "コース名"
    Next i
    
    Range("A1").Select
    
End Sub

Sub saveBook()
' 保存
    
    ActiveWorkbook.Save

End Sub

Sub toDisplaySheet()
' ランキングページへ移動
    
    Range("A1").Select
    Sheets("ランキング").Select
    Range("A1").Select
    
End Sub

Sub clearAllData()
' 全てのデータの削除
' デバッグ用

    ' Dataシート
    Sheets("Data").Select
    
    Cells(1, 9).value = "0"
    
    For rowNum = 2 To TRACK_NUM + 1
        Cells(rowNum, 2).value = "0"
        Cells(rowNum, 3).value = "0"
        Cells(rowNum, 4).value = "0"
    Next rowNum
    
    ' ランキングシート
    deleteRanks
    
    Range("A1").Select

End Sub

Sub inputSampleData()
' サンプルデータをセットする
' デバッグ用
    Dim rootPath, filePath
    Dim str As String

    ' サンプルデータファイルパス
    rootPath = ThisWorkbook.Path
    filePath = rootPath & "\sampleData\sampleData.txt"
    
    ' データを読み込む
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        str = .ReadText
        .Close
    End With
    
    ' クリップボードに値をセット
    Dim dataObj As DataObject
    Set dataObj = New DataObject
    dataObj.SetText str
    dataObj.PutInClipboard
    Set dataObj = Nothing
    
    ' Dataシートに貼り付け
    ' NOTE: カンマ区切りのデータを受け付けるように設定する必要がある
    Sheets("Data").Select
    Range("B2").Select
    ActiveSheet.Paste
    
    Range("A1").Select
    
End Sub
