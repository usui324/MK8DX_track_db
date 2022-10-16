Attribute VB_Name = "DisplayService"
Sub setRanks()
' 全てのランキングをセットする
    deleteRanks

    setRankPopular
    setRankNotPopular
    setRankGoodRank
    setRankBadRank
    setRankHighValueRank
    setRankGoodPointRank

    sortDataDefault

End Sub

Sub setRankPopular()
' 人気コースランキング(B2)の作成
    
    ' レース数_降順でソート
    Sheets("Data").Select
    sortDataByRaceNumDown
    
    ' 上から10列をランキングに表示
    Call setRankTopTen(2, 2, 4)
    
    Sheets("ランキング").Select
    Range("A1").Select

End Sub

Sub setRankNotPopular()
' 不人気コースランキング(B14)の作成
    
    ' レース数_昇順でソート
    Sheets("Data").Select
    sortDataByRaceNumUp
    
    ' 上から10列をランキングに表示
    Call setRankTopTen(14, 2, 4)
    
    Sheets("ランキング").Select
    Range("A1").Select
    
End Sub

Sub setRankGoodRank()
' 得意コース（平均順位）ランキング(F2)の作成
    
    ' 平均順位_昇順でソート
    Sheets("Data").Select
    sortDataByAvgRankUp
    
    ' 上から10列をランキングに表示
    Call setRankTopTen(2, 6, 5)
    
    ' ソートフィルターのリセット
    sortFilterReset
    
    Sheets("ランキング").Select
    Range("A1").Select

End Sub

Sub setRankBadRank()
' 不得意コースランキング(F14)の作成
    
    ' 平均順位_降順でソート
    Sheets("Data").Select
    sortDataByAvgRankDown
    
    ' 上から10列をランキングに表示
    Call setRankTopTen(14, 6, 5)
    
    ' ソートフィルターのリセット
    sortFilterReset
    
    Sheets("ランキング").Select
    Range("A1").Select

End Sub

Sub setRankHighValueRank()
' 上位期待値ランキング(J14)の作成
    
    ' 上位期待値_降順でソート
    Sheets("Data").Select
    sortDataByAvgValueUp
    
    ' 上から10列をランキングに表示
    Call setRankTopTen(14, 10, 7)
    
    ' ソートフィルターのリセット
    sortFilterReset
    
    Sheets("ランキング").Select
    Range("A1").Select

End Sub

Sub setRankGoodPointRank()
' 得意コース（平均得点）ランキング(J2)の作成
    
    ' 平均得点_降順でソート
    Sheets("Data").Select
    sortDataByAvgPointDown
    
    ' 上から10列をランキングに表示
    Call setRankTopTen(2, 10, 6)
    
    ' ソートフィルターのリセット
    sortFilterReset
    
    Sheets("ランキング").Select
    Range("A1").Select

End Sub

Function setRegurationRaceNum() As Integer
' ランキング掲載の基準レース数をセット
    
    setRegurationRaceNum = Sheets("ランキング").Cells(3, 14).value
End Function

Sub deleteRanks()
' 全てのランキングを削除する

    Sheets("ランキング").Select
    
    Range("C3", "D12").Select
    Selection.ClearContents
    
    Range("C15", "D24").Select
    Selection.ClearContents

    Range("G3", "H12").Select
    Selection.ClearContents
    
    Range("G15", "H24").Select
    Selection.ClearContents
    
    Range("K3", "L12").Select
    Selection.ClearContents
    
    Range("K15", "L24").Select
    Selection.ClearContents
End Sub











