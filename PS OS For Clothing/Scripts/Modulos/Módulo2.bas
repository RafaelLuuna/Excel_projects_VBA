Attribute VB_Name = "Módulo2"
Dim Clientes As Collection
Dim BestID As Integer

Private Sub AcharMaisFrequente(Matriz As Collection)
    Dim i, iCol As Integer
    Dim Score(9999) As Integer
    BestID = 1
    For iCol = 1 To Matriz.Count
        For i = 2 To Planilha4.Range("H:H").Cells.SpecialCells(xlCellTypeConstants).Count
            If Planilha4.Cells(i, 8).Value = Matriz.Item(iCol) Then Score(iCol) = Score(iCol) + 1
        Next
        If Score(iCol) >= Score(BestID) Then BestID = iCol
    Next
End Sub
