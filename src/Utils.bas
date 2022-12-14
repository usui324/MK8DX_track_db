Attribute VB_Name = "Utils"
Sub checkInput()
' 入力内容チェック

    Sheets(SHEET1_NAME).Select
    
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

    Sheets(SHEET1_NAME).Select
    
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
    Sheets(SHEET2_NAME).Select
    Range("A1").Select
    
End Sub

Sub exportData()
' データをtxtファイルにエクスポート
    ' エクスポートファイルを指定
    ChDir ThisWorkbook.Path
    Dim saveFileName As String
    saveFileName = Application.GetSaveAsFilename(InitialFileName:="trackData.txt", filefilter:="コースデータ,*.txt")

    ' キャンセル処理
    If saveFileName = "False" Then
        Exit Sub
    End If

    ' 出力する対象シート
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET3_NAME)
    
    ' ファイルに書き込み
    Open saveFileName For Output As #1
    ' 模擬回数
    Print #1, ws.Cells(1, 8).value & "," & ws.Cells(1, 9).value
    ' trackData
    Dim i As Integer
    For i = 1 To TRACK_NUM
        Print #1, printLine(ws, i + 1)
    Next i
    
    Close #1
    
    MsgBox saveFileName & "にデータを出力しました", vbInformation

End Sub

Function printLine(ws As Worksheet, rowNo As Integer) As String
' ファイル出力する一行の文字列を返す
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
' txtファイルからデータをインポート
    Dim openFileName As String
    Dim ws As Worksheet
    Dim line As String
    Dim arr As Variant
    Dim rowNo As Integer
    Dim columnNo As Integer

    ' インポートファイルを指定
    ChDir ThisWorkbook.Path
    openFileName = Application.GetOpenFilename("コースデータ,*.txt", , "インポートするデータファイルを指定")
    
    ' キャンセル処理
    If openFileName = "False" Then
        Exit Sub
    End If
    
    ' 入力対象シート
    Set ws = ThisWorkbook.Worksheets(SHEET3_NAME)
    
    ' 模擬回数の入力
    Open openFileName For Input As #1
    Line Input #1, line
    arr = Split(line, ",")
    ws.Cells(1, 9).value = arr(1)
    
    ' コースデータの入力
    rowNo = 2
    While Not EOF(1)
        Line Input #1, line
        arr = Split(line, ",")
        
        ' 配列長が5以上の場合はエラー
        If UBound(arr) >= 4 Then
            MsgBox "データが不正です", vbExclamation
            Exit Sub
        End If
        
        For columnNo = LBound(arr) To UBound(arr)
            ws.Cells(rowNo, columnNo + 1).value = arr(columnNo)
        Next columnNo
        rowNo = rowNo + 1
    Wend
    Close #1
    
    ' ランキングデータを更新
    setRanks
    
    MsgBox "データをインポートしました", vbInformation
End Sub

Sub setWindowSizeDataInput()
' Windowサイズの調整
    Application.WindowState = xlNormal
    ActiveWindow.Width = 430
    ActiveWindow.Height = 720

End Sub

Sub setWindowSizeRanking()
' Windowサイズの調整
    Application.WindowState = xlNormal
    ActiveWindow.Width = 1100
    ActiveWindow.Height = 720
    
    ' 画面外に行かないよう位置も調整する
    ActiveWindow.Top = 0
    ActiveWindow.Left = 0

End Sub

Sub setSheet()
' シートをセット
    Sheets(SHEET1_NAME).Select

End Sub
