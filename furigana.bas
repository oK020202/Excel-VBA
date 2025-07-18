Attribute VB_Name = "furigana"
Sub ボタン1_Click()
    Call furigana
End Sub
Private Sub furigana()
    '振り仮名を自動で振る
    Dim st As Worksheet
    Dim Rng As Range
    Const offSet As Long = 3
    Const contNum As Long = 12
    Dim i As Long
    Set st = ThisWorkbook.Sheets("テスト名簿")
    For i = offSet + 1 To contNum + offSet
        Set Rng = st.Cells(i, 3)
        st.Cells(i, 4).Value = WorksheetFunction.Phonetic(Rng)
    Next i
End Sub
