Attribute VB_Name = "Liblary"
Sub ExportAll()
' モジュールを全てエクスポートする

    ' モジュール
    Dim module As VBComponent
    Dim moduleList As VBComponents
    
    ' 拡張子
    Dim extension
    ' 対象ブックのパス
    Dim targetPath
    ' エクスポートファイルパス
    Dim exportPath
    ' 対象ブックオブジェクト
    Dim targetBook


    ' このブックを対象とする
    Set targetBook = ThisWorkbook
    targetPath = targetBook.Path

    ' モジュール一覧を取得
    Set moduleList = targetBook.VBProject.VBComponents
    
    ' 各モジュールに対する処理
    For Each module In moduleList
        ' クラス
        If module.Type = vbext_ct_ClassModule Then
            extension = "cls"
        ' フォーム
        ElseIf module.Type = vbext_ct_MSForm Then
            extension = "frm"
        ' 標準モジュール
        ElseIf module.Type = vbext_ct_StdModule Then
            extension = "bas"
        ' その他
        Else
            MsgBox module.Type
            GoTo CONTINUE
        End If
        
        ' エクスポート処理
        exportPath = targetPath & "\src\" & module.Name & "." & extension
        Call module.Export(exportPath)
        
        ' 出力先確認用ログ
        Debug.Print exportPath
        
CONTINUE:
    Next

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

