Attribute VB_Name = "DisplayService"
Sub setRanks()
' �S�Ẵ����L���O���Z�b�g����
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
' �l�C�R�[�X�����L���O(B2)�̍쐬
    
    ' ���[�X��_�~���Ń\�[�g
    Sheets(SHEET3_NAME).Select
    sortDataByRaceNumDown
    
    ' �ォ��10��������L���O�ɕ\��
    Call setRankTopTen(2, 2, 4)
    
    Sheets(SHEET2_NAME).Select
    Range("A1").Select

End Sub

Sub setRankNotPopular()
' �s�l�C�R�[�X�����L���O(B14)�̍쐬
    
    ' ���[�X��_�����Ń\�[�g
    Sheets(SHEET3_NAME).Select
    sortDataByRaceNumUp
    
    ' �ォ��10��������L���O�ɕ\��
    Call setRankTopTen(14, 2, 4)
    
    Sheets(SHEET2_NAME).Select
    Range("A1").Select
    
End Sub

Sub setRankGoodRank()
' ���ӃR�[�X�i���Ϗ��ʁj�����L���O(F2)�̍쐬
    
    ' ���Ϗ���_�����Ń\�[�g
    Sheets(SHEET3_NAME).Select
    sortDataByAvgRankUp
    
    ' �ォ��10��������L���O�ɕ\��
    Call setRankTopTen(2, 6, 5)
    
    ' �\�[�g�t�B���^�[�̃��Z�b�g
    sortFilterReset
    
    Sheets(SHEET2_NAME).Select
    Range("A1").Select

End Sub

Sub setRankBadRank()
' �s���ӃR�[�X�����L���O(F14)�̍쐬
    
    ' ���Ϗ���_�~���Ń\�[�g
    Sheets(SHEET3_NAME).Select
    sortDataByAvgRankDown
    
    ' �ォ��10��������L���O�ɕ\��
    Call setRankTopTen(14, 6, 5)
    
    ' �\�[�g�t�B���^�[�̃��Z�b�g
    sortFilterReset
    
    Sheets(SHEET2_NAME).Select
    Range("A1").Select

End Sub

Sub setRankHighValueRank()
' ��ʊ��Ғl�����L���O(J14)�̍쐬
    
    ' ��ʊ��Ғl_�~���Ń\�[�g
    Sheets(SHEET3_NAME).Select
    sortDataByAvgValueUp
    
    ' �ォ��10��������L���O�ɕ\��
    Call setRankTopTen(14, 10, 7)
    
    ' �\�[�g�t�B���^�[�̃��Z�b�g
    sortFilterReset
    
    Sheets(SHEET2_NAME).Select
    Range("A1").Select

End Sub

Sub setRankGoodPointRank()
' ���ӃR�[�X�i���ϓ��_�j�����L���O(J2)�̍쐬
    
    ' ���ϓ��__�~���Ń\�[�g
    Sheets(SHEET3_NAME).Select
    sortDataByAvgPointDown
    
    ' �ォ��10��������L���O�ɕ\��
    Call setRankTopTen(2, 10, 6)
    
    ' �\�[�g�t�B���^�[�̃��Z�b�g
    sortFilterReset
    
    Sheets(SHEET2_NAME).Select
    Range("A1").Select

End Sub

Function setRegurationRaceNum() As Integer
' �����L���O�f�ڂ̊���[�X�����Z�b�g
    
    setRegurationRaceNum = Sheets(SHEET2_NAME).Cells(3, 14).value
End Function

Sub deleteRanks()
' �S�Ẵ����L���O���폜����

    Sheets(SHEET2_NAME).Select
    
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











