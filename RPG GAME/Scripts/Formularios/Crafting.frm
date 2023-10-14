VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Crafting 
   Caption         =   "Crafting"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5865
   OleObjectBlob   =   "Crafting.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Crafting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ItemID As String
Private Needs(6) As Item

Private Sub CommandButton1_Click()
    Call Craft
End Sub

Private Sub CraftList_Change()
    Dim i As Integer
    For i = 0 To CraftList.ListCount - 1
        If CraftList.Selected(i) = True Then
            ItemID = Left(CraftList.List(i), InStr(1, CraftList.List(i), "(") - 2)
            ItemImg.Picture = LoadPicture(Application.ThisWorkbook.Path & "\texture\item\" & ItemID & ".gif")
        End If
    Next
    
    Needs(1).ID = ""
    Needs(1).Qnt = 0
    Needs(2).ID = ""
    Needs(2).Qnt = 0
    Needs(3).ID = ""
    Needs(3).Qnt = 0
    Needs(4).ID = ""
    Needs(4).Qnt = 0
    Needs(5).ID = ""
    Needs(5).Qnt = 0
    Needs(6).ID = ""
    Needs(6).Qnt = 0
    
    lbNeed1.Caption = ""
    lbNeed2.Caption = ""
    lbNeed3.Caption = ""
    lbNeed4.Caption = ""
    lbNeed5.Caption = ""
    lbNeed6.Caption = ""
    
    lbHave1.Caption = ""
    lbHave2.Caption = ""
    lbHave3.Caption = ""
    lbHave4.Caption = ""
    lbHave5.Caption = ""
    lbHave6.Caption = ""
    
    Select Case ItemID
        Case "WoodPlate"
            lbNeed1.Caption = "3 Woods"
            lbHave1.Caption = InventoryFunctions.CountItem(1, "Wood") & " Woods"
            Needs(1).ID = "Wood"
            Needs(1).Qnt = 3
        Case "Chest"
            lbNeed1.Caption = "6 WoodPlate"
            lbHave1.Caption = InventoryFunctions.CountItem(1, "WoodPlate") & " WoodPlate"
            Needs(1).ID = "WoodPlate"
            Needs(1).Qnt = 6
        Case "CraftTable"
            lbNeed1.Caption = "1 WoodPlate"
            lbHave1.Caption = InventoryFunctions.CountItem(1, "WoodPlate") & " WoodPlate"
            Needs(1).ID = "WoodPlate"
            Needs(1).Qnt = 1
            lbNeed2.Caption = "4 Woods"
            lbHave2.Caption = InventoryFunctions.CountItem(1, "Wood") & " Woods"
            Needs(2).ID = "Wood"
            Needs(2).Qnt = 4
        Case "Stone_Axe"
            lbNeed1.Caption = "1 Wood"
            lbHave1.Caption = InventoryFunctions.CountItem(1, "Wood") & " Woods"
            Needs(1).ID = "Wood"
            Needs(1).Qnt = 1
            lbNeed2.Caption = "1 Rock"
            lbHave2.Caption = InventoryFunctions.CountItem(1, "Rock") & " Rock"
            Needs(2).ID = "Rock"
            Needs(2).Qnt = 1
    End Select
End Sub


Private Sub CraftList_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Select Case KeyCode
    Case 87, 38
    Case 65, 37
    Case 83, 40
    Case 68, 39
    Case 69, 67
        Unload Me
    Case 70, 90, 13
        Call Craft
    Case Else
        KeyCode = 0
End Select
End Sub

Private Sub Craft()
    Dim i As Integer
    Dim CraftedItem As Item
    Dim CraftSlot As Integer
    
    For i = 0 To CraftList.ListCount
        If CraftList.Selected(i) = True Then
            CraftedItem.ID = Left(CraftList.List(i), InStr(1, CraftList.List(i), "(") - 2)
        End If
    Next
    
    For i = 1 To 6
        If Not Needs(i).ID = "Null" Then
            If InventoryFunctions.CountItem(1, Needs(i).ID) < Needs(i).Qnt Then
                MsgBox "you don't have enough resources to craft this item."
                Exit Sub
            End If
        End If
    Next
    
    Select Case CraftedItem.ID
        Case "Stone_Axe"
            If InventoryFunctions.FindItem(1, "Null") = 0 Then
                MsgBox "Your inventory is full."
                Exit Sub
            Else
                CraftSlot = InventoryFunctions.FindItem(1, "Null")
            End If
        Case Else
            If InventoryFunctions.FindItem(1, CraftedItem.ID) = 0 Then
            If InventoryFunctions.FindItem(1, "Null") = 0 Then
                MsgBox "Your inventory is full."
                Exit Sub
            Else
                CraftSlot = InventoryFunctions.FindItem(1, "Null")
            End If
            Else
                CraftSlot = InventoryFunctions.FindItem(1, CraftedItem.ID)
            End If
    End Select
    
    Select Case CraftedItem.ID
        Case "WoodPlate"
            Call InventoryFunctions.ChangeSlot(1, InventoryFunctions.FindItem(1, "Wood"), , DATA.InventoryArray(1).InventorySlots(InventoryFunctions.FindItem(1, "Wood")).Qnt - 3, 1)
            Call InventoryFunctions.ChangeSlot(1, CraftSlot, "WoodPlate", DATA.InventoryArray(1).InventorySlots(CraftSlot).Qnt + 1, 1)
        Case "Chest"
            Call InventoryFunctions.ChangeSlot(1, InventoryFunctions.FindItem(1, "WoodPlate"), , DATA.InventoryArray(1).InventorySlots(InventoryFunctions.FindItem(1, "WoodPlate")).Qnt - 6, 1)
            Call InventoryFunctions.ChangeSlot(1, CraftSlot, "Chest", DATA.InventoryArray(1).InventorySlots(CraftSlot).Qnt + 1, 1)
        Case "CraftTable"
            Call InventoryFunctions.ChangeSlot(1, InventoryFunctions.FindItem(1, "WoodPlate"), , DATA.InventoryArray(1).InventorySlots(InventoryFunctions.FindItem(1, "WoodPlate")).Qnt - 1, 1)
            Call InventoryFunctions.ChangeSlot(1, InventoryFunctions.FindItem(1, "Wood"), , DATA.InventoryArray(1).InventorySlots(InventoryFunctions.FindItem(1, "Wood")).Qnt - 4, 1)
            Call InventoryFunctions.ChangeSlot(1, CraftSlot, "CraftTable", DATA.InventoryArray(1).InventorySlots(CraftSlot).Qnt + 1, 1)
        Case "Stone_Axe"
            Call InventoryFunctions.ChangeSlot(1, InventoryFunctions.FindItem(1, "Rock"), , DATA.InventoryArray(1).InventorySlots(InventoryFunctions.FindItem(1, "Rock")).Qnt - 1, 1)
            Call InventoryFunctions.ChangeSlot(1, InventoryFunctions.FindItem(1, "Wood"), , DATA.InventoryArray(1).InventorySlots(InventoryFunctions.FindItem(1, "Wood")).Qnt - 1, 1)
            Call InventoryFunctions.ChangeSlot(1, CraftSlot, "Stone_Axe", 1, 7)
    End Select
    
    Call Update
    
End Sub

Private Sub Update()
Dim i, x
x = 0
    For i = 0 To CraftList.ListCount - 1
        If CraftList.Selected(i) = True Then x = i
    Next
CraftList.Clear
CraftList.AddItem "WoodPlate (" & InventoryFunctions.CountItem(1, "WoodPlate") & ")"
CraftList.AddItem "Chest (" & InventoryFunctions.CountItem(1, "Chest") & ")"
CraftList.AddItem "CraftTable (" & InventoryFunctions.CountItem(1, "CraftTable") & ")"
CraftList.AddItem "Stone_Axe (" & InventoryFunctions.CountItem(1, "Stone_Axe") & ")"
CraftList.Selected(x) = True
End Sub

Private Sub UserForm_Initialize()
Call Update
End Sub
