Attribute VB_Name = "InventoryFunctions"
Public InventoryFullTest As Boolean

'Private Sub Test()
'    MsgBox WorksheetFunction.VLookup("Apple", ItemStats.Range("A:C"), 10, False)
'End Sub

Public Function CheckItemStats(ItemID As String, Stats As Integer) As String
    On Error Resume Next
    CheckItemStats = WorksheetFunction.VLookup(ItemID, ItemStats.Range("A:C"), Stats + 1, False)
    If Err.Number = 1004 Then CheckItemStats = "##ERROR"
End Function

Public Sub AddItem(InventoryID As Integer, ItemID As String, ItemQnt As Integer, ItemDurabillity As Integer)
    InventoryFullTest = False
    Dim NewItem As Item
    Dim Slot As Integer
    
    
    NewItem.ID = ItemID
    NewItem.Qnt = ItemQnt
    NewItem.Durabillity = ItemDurabillity
    
    If NewItem.Qnt <= 0 Or NewItem.Durabillity <= 0 Then Exit Sub
    
    If InventoryFunctions.FindItem(InventoryID, "Null") = 0 Then
        Select Case WorksheetFunction.VLookup(ItemID, ItemStats.Range("A:B"), 2, False)
            Case "s"
                If CountItem(InventoryID, ItemID) = 0 Then
                    InventoryFullTest = True
                    Exit Sub
                Else
                    Slot = FindItem(InventoryID, ItemID)
                    Call ChangeSlot(InventoryID, Slot, ItemID, ItemQnt + DATA.InventoryArray(InventoryID).InventorySlots(Slot).Qnt, ItemDurabillity)
                    Exit Sub
                End If
            Case "n"
                InventoryFullTest = True
                Exit Sub
        End Select
    End If
    
    Slot = FindItem(InventoryID, "Null")
    
    InventoryData.Cells(Slot + 2, InventoryID) = NewItem.ID & "," & NewItem.Qnt & "," & NewItem.Durabillity
    DATA.InventoryArray(InventoryID).InventorySlots(Slot) = NewItem
    
    
End Sub

Public Sub ChangeSlot(InventoryID As Integer, Slot As Integer, Optional ItemID As String = "Omitido", Optional ItemQnt = "Omitido", Optional ItemDurabillity = "Omitido")
    Dim NewItem As Item
    
    If Slot > DATA.InventoryArray(InventoryID).InventorySize Then Exit Sub
    
    If ItemID = "Omitido" Then ItemID = DATA.InventoryArray(InventoryID).InventorySlots(Slot).ID
    NewItem.ID = ItemID
    
    If ItemQnt = "Omitido" Then ItemQnt = DATA.InventoryArray(InventoryID).InventorySlots(Slot).Qnt
    NewItem.Qnt = ItemQnt
    
    If ItemDurabillity = "Omitido" Then ItemDurabillity = DATA.InventoryArray(InventoryID).InventorySlots(Slot).Durabillity
    NewItem.Durabillity = ItemDurabillity
    
    If NewItem.Qnt <= 0 Or NewItem.Durabillity <= 0 Then
        NewItem.ID = "Null"
        NewItem.Qnt = 0
        NewItem.Durabillity = 0
    End If
    
    InventoryData.Cells(Slot + 2, InventoryID) = NewItem.ID & "," & NewItem.Qnt & "," & NewItem.Durabillity
    DATA.InventoryArray(InventoryID).InventorySlots(Slot) = NewItem
End Sub

Public Function FindItem(InventoryID As Integer, ItemID As String, Optional Start As Integer = 1) As Integer
    Dim i As Integer
    FindItem = 0
    For i = Start To DATA.InventoryArray(InventoryID).InventorySize
        If DATA.InventoryArray(InventoryID).InventorySlots(i).ID = ItemID Then
            FindItem = i
            Exit Function
        End If
    Next
    For i = Start To DATA.InventoryArray(InventoryID).InventorySize
        If DATA.InventoryArray(InventoryID).InventorySlots(i).ID = "Null" Then
            FindItem = i
            Exit Function
        End If
    Next
End Function

Public Function CountItem(InventoryID As Integer, ItemID As String) As Integer
    Dim i As Integer
    Dim Count As Integer
    
    For i = 1 To DATA.InventoryArray(InventoryID).InventorySize
        If ItemID = DATA.InventoryArray(InventoryID).InventorySlots(i).ID Then Count = Count + DATA.InventoryArray(InventoryID).InventorySlots(i).Qnt
    Next
    
    CountItem = Count
    
End Function

Public Function CountWeapons() As Integer
Dim i, iWp As Integer
For i = 2 To WpData.Range("A:A").SpecialCells(xlCellTypeConstants).Count
    If InventoryFunctions.CountItem(1, WpData.Cells(i, 1)) > 0 Then
        For iWp = 1 To InventoryFunctions.CountItem(1, WpData.Cells(i, 1))
            CountWeapons = CountWeapons + 1
        Next
    End If
Next
End Function
