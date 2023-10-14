Attribute VB_Name = "Módulo1"

Public Function SomarCol(ByRef Col As Collection) As Integer
    Dim i, Value As Integer
    For i = 1 To Col.Count
        If Not Col.Item(i).Text = "" Then Value = Value + CInt(Col.Item(i).Text)
    Next
    SomarCol = Value
End Function

