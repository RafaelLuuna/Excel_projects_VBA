Attribute VB_Name = "ChestNames"
Public Function Find(xCoord As Integer, yCoord As Integer, Scene As String) As String
Dim StrCoord As String
StrCoord = xCoord & "," & yCoord

Select Case Scene
    Case "Casa_2"
        Select Case StrCoord
            Case "5,6"
                Find = "ChestHouse1"
            Case "6,6"
                Find = "ChestHouse2"
        End Select
    Case "Mundo_2"
        Select Case StrCoord
            Case "3,9"
                Find = "ChestWorld1"
            Case "4,4"
                Find = "ChestWorld2"
            Case "5,19"
                Find = "ChestWorld3"
            Case "19,11"
                Find = "ChestWorld4"
        End Select
End Select
End Function
