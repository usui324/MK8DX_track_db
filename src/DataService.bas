Attribute VB_Name = "DataService"
Sub sortDataByRaceNumDown()
Attribute sortDataByRaceNumDown.VB_ProcData.VB_Invoke_Func = " \n14"
' レース数の降順でソートを実施する

    Sheets("Data").Select

    Columns("A:F").Select
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add2 Key:=Range("D2:D97"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").Sort
        .SetRange Range("A1:F97")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1").Select
    
End Sub

Sub sortDataByRaceNumUp()
Attribute sortDataByRaceNumUp.VB_ProcData.VB_Invoke_Func = " \n14"
' レース数の昇順でソートを実施する

    Sheets("Data").Select
    
    Columns("A:F").Select
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add2 Key:=Range("D2:D97"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").Sort
        .SetRange Range("A1:F97")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1").Select
    
End Sub

Sub sortDataDefault()
Attribute sortDataDefault.VB_ProcData.VB_Invoke_Func = " \n14"
' デフォルト順にソート

    Sheets("Data").Select
    
    Columns("A:F").Select
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").Sort.SortFields.Add2 Key:=Range("A2:A97"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, _
        customOrder:=TRACK_LIST_STR, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").Sort
        .SetRange Range("A1:F97")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1").Select
    
End Sub

Sub sortFilterReset()
'ソート・フィルターのリセット
    
    Sheets("Data").Select
    Selection.AutoFilter
    
End Sub

Sub sortDataByAvgRankUp()
Attribute sortDataByAvgRankUp.VB_ProcData.VB_Invoke_Func = " \n14"
' 平均順位の昇順にソート

    Dim regurationRaceNum As Integer
    regurationRaceNum = setRegurationRaceNum

    Sheets("Data").Select
    
    Columns("A:F").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$F$97").AutoFilter Field:=4, Criteria1:=">=" & regurationRaceNum, _
    Operator:=xlAnd
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "E1:E97"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    Range("A1").Select
    
End Sub

Sub sortDataByAvgRankDown()
Attribute sortDataByAvgRankDown.VB_ProcData.VB_Invoke_Func = " \n14"
' 平均順位の降順にソート
    
    Dim regurationRaceNum As Integer
    regurationRaceNum = setRegurationRaceNum
    
    Sheets("Data").Select
    
    Columns("A:F").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$F$97").AutoFilter Field:=4, Criteria1:=">=" & regurationRaceNum, _
    Operator:=xlAnd
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "E1:E97"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1").Select
    
End Sub

Sub sortDataByAvgValueUp()
' 上位期待値の降順にソート

    Dim regurationRaceNum As Integer
    regurationRaceNum = setRegurationRaceNum

    Sheets("Data").Select
    
    Columns("A:G").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$G$97").AutoFilter Field:=4, Criteria1:=">=" & regurationRaceNum, _
    Operator:=xlAnd
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "G1:G97"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1").Select
    
End Sub

Sub sortDataByAvgPointDown()
' 平均得点の降順にソート

    Dim regurationRaceNum As Integer
    regurationRaceNum = setRegurationRaceNum

    Sheets("Data").Select
    
    Columns("A:F").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$F$97").AutoFilter Field:=4, Criteria1:=">=" & regurationRaceNum, _
    Operator:=xlAnd
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Data").AutoFilter.Sort.SortFields.Add2 Key:=Range( _
        "F1:F97"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Data").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A1").Select
    
End Sub

Sub setRankTopTen(titleRow As Integer, titleColumn As Integer, valueColumn As Integer)
' データシートの上位10列をランキングシートの指定位置にコピー
    
    Dim trackName As String
    Dim value As Double
    Dim endRow As Integer
    Dim dataRow As Integer
    Dim displayRow As Integer
    endRow = 11
    dataRow = 2
    displayRow = 1
    
    
    While dataRow <= endRow
        If Sheets("Data").Rows(dataRow).Hidden Then
            endRow = endRow + 1
            GoTo CONTINUE:
        End If
        
        Sheets("Data").Select
        trackName = Cells(dataRow, 1).Text
        value = Val(Cells(dataRow, valueColumn).Text)
    
        Sheets("ランキング").Select
        Cells(titleRow + displayRow, titleColumn + 1).value = trackName
        Cells(titleRow + displayRow, titleColumn + 2).value = value
        displayRow = displayRow + 1
    
CONTINUE:
        dataRow = dataRow + 1

         ' コース数を超えたら終了
        If dataRow > TRACK_NUM + 1 Then
            GoTo BREAK:
        End If

    Wend
BREAK:
    
    Sheets("Data").Select
    Range("A1").Select

End Sub
