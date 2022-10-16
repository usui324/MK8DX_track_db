Attribute VB_Name = "ResistService"
' データ登録
Sub registInput()

' 各コースに対して同じメソッドを実行
For i = 3 To 14

    ' コース名・順位未入力列はスキップ
    Sheets("データ入力").Select
    
    If Cells(i, 2).value = "コース名" Or Cells(i, 3).value = "" Then
        GoTo CONTINUE1:
        
    End If
    
    Dim trackName As String
    trackName = Cells(i, 2).value
    Dim rankValue As Integer
    rankValue = Cells(i, 3).value
    
    ' 順位に対する得点を計算
    Sheets("MasterData").Select
    Dim pointValue As Integer
    pointValue = Cells(rankValue + 1, 4).value

    ' Dataからコース名に合致する列を探索
    Dim trackRow As Integer
    
    Sheets("Data").Select
    For j = 2 To 97
        If Cells(j, 1).value = trackName Then
            trackRow = j
            GoTo CONTINUE2:
        End If
CONTINUE2:
    Next j
    
    ' データの書き込み
    Sheets("Data").Select
    Cells(trackRow, 2).value = Cells(trackRow, 2).value + rankValue
    Cells(trackRow, 3).value = Cells(trackRow, 3).value + pointValue
    Cells(trackRow, 4).value = Cells(trackRow, 4).value + 1
    
CONTINUE1:
Next i

' 模擬回数の加算
Sheets("Data").Select
Cells(1, 9).value = Cells(1, 9).value + 1

Sheets("データ入力").Select
Range("A1").Select

End Sub

