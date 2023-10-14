VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfrmInventory 
   Caption         =   "UserForm1"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11550
   OleObjectBlob   =   "usfrmInventory.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfrmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SelectedRow As Integer
Private OptionMode As Boolean
Private SelectMode As Boolean
Private SelectModeSlot As Integer
Private SelectedOption As String
Private CombA, CombB
Private SlotA As Integer
Public SelectedSlot As Integer
Public InChest As Boolean
Public IsEmpty As Boolean
Public Usable As Boolean
Public InventoryID As Integer


Private Sub UserForm_Initialize()
    SelectedSlot = 1
    OptionMode = False
    SelectMode = False
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Dim InventorySize As Integer
Dim oControl As Control

Select Case OptionMode
    Case True
        
        Select Case KeyCode
            Case 87, 38 'W
                Me.Controls.Item("Options").Selected(SelectedRow) = False
                If Not SelectedRow = 0 Then SelectedRow = SelectedRow - 1
                Me.Controls.Item("Options").Selected(SelectedRow) = True
            Case 83, 40 'S
                Me.Controls.Item("Options").Selected(SelectedRow) = False
                If Not SelectedRow = Me.Controls.Item("Options").ListCount - 1 Then SelectedRow = SelectedRow + 1
                Me.Controls.Item("Options").Selected(SelectedRow) = True
            Case 69, 67 'E
                Me.Controls.Clear
                Unload Me
            Case 32, 70, 13, 90 'Enter
                Dim SelectedItem As Item
                Dim i As Integer
                SelectedItem = DATA.InventoryArray(InventoryID).InventorySlots(SelectedSlot)
                
                For i = 0 To Me.Controls.Item("Options").ListCount - 1
                    If Me.Controls.Item("Options").Selected(i) = True Then SelectedOption = Me.Controls.Item("Options").List(i)
                Next
                
                                     
                Select Case SelectedOption
                    Case "Catch"
                        If InventoryFunctions.FindItem(1, SelectedItem.ID) = 0 Then
                            MsgBox "Inventário Cheio!"
                        Else
                            Select Case SelectedItem.ID
                                Case "Stone_Axe"
                                    Call InventoryFunctions.ChangeSlot(1, InventoryFunctions.FindItem(1, "Null"), SelectedItem.ID, SelectedItem.Qnt, SelectedItem.Durabillity)
                                Case Else
                                    Call InventoryFunctions.ChangeSlot(1, InventoryFunctions.FindItem(1, SelectedItem.ID), SelectedItem.ID, DATA.InventoryArray(1).InventorySlots(InventoryFunctions.FindItem(1, SelectedItem.ID)).Qnt + SelectedItem.Qnt, SelectedItem.Durabillity)
                            End Select
                            Call InventoryFunctions.ChangeSlot(InventoryID, SelectedSlot, "Null", 0, 0)
                        End If
                    Case "Delete"
                        If MsgBox("Are you sure? You can't undo this action.", vbYesNo, "Are you sure?") = vbYes Then Call InventoryFunctions.ChangeSlot(InventoryID, SelectedSlot, "Null", 0, 0)
                    Case "Cancel Comb."
                        SlotA = 0
                        CombA = ""
                    Case "Combine"
                        If CombA = "" Then
                            Me.Controls("Slot" & SelectedSlot).BackColor = &HC0C0C0
                            CombA = DATA.InventoryArray(InventoryID).InventorySlots(SelectedSlot).ID
                            SlotA = SelectedSlot
                        Else
                            If Not SelectedSlot = SlotA Then
                                CombB = DATA.InventoryArray(InventoryID).InventorySlots(SelectedSlot).ID
                                If CombA = CombB Then
                                    Dim CombA_Test As Boolean
                                    CombA_Test = False
                                    If WorksheetFunction.Match(CombA, WpData.Range("A:A")) > 0 Then CombA_Test = True
                                    Select Case CombA_Test
                                        Case True
                                            MsgBox Prompt:="You can't stack that type of item.", Buttons:=vbInformation, Title:="Erro ao combinar itens"
                                        Case False
                                            Call InventoryFunctions.ChangeSlot(InventoryID, SelectedSlot, , DATA.InventoryArray(InventoryID).InventorySlots(SelectedSlot).Qnt + DATA.InventoryArray(InventoryID).InventorySlots(SlotA).Qnt, 1)
                                            Call InventoryFunctions.ChangeSlot(InventoryID, SlotA, "Null", 0, 0)
                                    End Select
                                Else
                                    MsgBox Prompt:="you can't combine these items", Buttons:=vbInformation, Title:="Erro ao combinar itens"
                                End If
                                CombA = ""
                                CombB = ""
                                SlotA = 0
                            End If
                        End If
                    Case "Place Item"
                        Call Load.InventorySlots(1)
                        OptionMode = False
                        SelectMode = True
                        SelectModeSlot = SelectedSlot
                    Case "Select"
                        usfrmQnt.MaxValue = DATA.InventoryArray(1).InventorySlots(SelectedSlot).Qnt
                        usfrmQnt.Show
                        Call InventoryFunctions.ChangeSlot(InventoryID, SelectModeSlot, DATA.InventoryArray(1).InventorySlots(SelectedSlot).ID, DATA.VarQnt, DATA.InventoryArray(1).InventorySlots(SelectedSlot).Durabillity)
                        If DATA.VarQnt = DATA.InventoryArray(1).InventorySlots(SelectedSlot).Qnt Then
                            Call InventoryFunctions.ChangeSlot(1, SelectedSlot, "Null", 0, 0)
                        Else
                            Call InventoryFunctions.ChangeSlot(1, SelectedSlot, , DATA.InventoryArray(1).InventorySlots(SelectedSlot).Qnt - DATA.VarQnt)
                        End If
                        SelectedSlot = SelectModeSlot
                        SelectMode = False
                    Case "Cancel"
                    If SelectMode Then
                        Me.Controls.Remove ("Options")
                        OptionMode = False
                    End If
                End Select
                
                If Not SelectMode Then
                    Me.Controls.Remove ("Options")
                    Call Load.InventorySlots(InventoryID)
                    OptionMode = False
                    If Not SlotA = 0 Then Me.Controls("Slot" & SlotA).BackColor = &HC0C0C0
                End If
        End Select
        
        
    Case False
        InventorySize = Me.Controls.Count / 2
        Select Case KeyCode
            Case 87, 38 'W
                If Not SelectedSlot = SlotA Then Me.Controls.Item("Slot" & SelectedSlot).BackColor = &HFFFFFF
                If SelectedSlot > InventoryDisplaySize Then SelectedSlot = SelectedSlot - InventoryDisplaySize
                Me.Controls.Item("Slot" & SelectedSlot).BackColor = &HE0E0E0
            Case 65, 37 'A
                If Not SelectedSlot = SlotA Then Me.Controls.Item("Slot" & SelectedSlot).BackColor = &HFFFFFF
                If Not SelectedSlot = 1 Then SelectedSlot = SelectedSlot - 1
                Me.Controls.Item("Slot" & SelectedSlot).BackColor = &HE0E0E0
            Case 83, 40 'S
                If Not SelectedSlot = SlotA Then Me.Controls.Item("Slot" & SelectedSlot).BackColor = &HFFFFFF
                If SelectedSlot <= InventorySize - InventoryDisplaySize Then SelectedSlot = SelectedSlot + InventoryDisplaySize
                Me.Controls.Item("Slot" & SelectedSlot).BackColor = &HE0E0E0
            Case 68, 39 'D
                If Not SelectedSlot = SlotA Then Me.Controls.Item("Slot" & SelectedSlot).BackColor = &HFFFFFF
                If Not SelectedSlot = InventorySize Then SelectedSlot = SelectedSlot + 1
                Me.Controls.Item("Slot" & SelectedSlot).BackColor = &HE0E0E0
            Case 69, 67 'E
                Me.Controls.Clear
                Unload Me
            Case 32, 70, 13, 90 'Enter
                Set oControl = Me.Controls.Add("Forms.ListBox.1", "Options")
                Select Case InventoryFunctions.CheckItemStats(DATA.InventoryArray(InventoryID).InventorySlots(SelectedSlot).ID, 2)
                    Case "s"
                        Usable = True
                    Case "n"
                        Usable = False
                End Select
                With oControl
                    .Height = Me.Controls.Item("Slot" & SelectedSlot).Height
                    .Width = Me.Controls.Item("Slot" & SelectedSlot).Width
                    .Top = Me.Controls.Item("Slot" & SelectedSlot).Top
                    .Left = Me.Controls.Item("Slot" & SelectedSlot).Left
                    .Enabled = False
                    If DATA.InventoryArray(InventoryID).InventorySlots(SelectedSlot).ID = "Null" Then
                            IsEmpty = True
                    Else:   IsEmpty = False
                    End If
                    If Not CombA = "" Then
                        .AddItem "Combine"
                        .AddItem "Cancel Comb."
                        .AddItem "Cancel"
                    ElseIf SelectMode Then
                        .AddItem "Select"
                        .AddItem "Cancel"
                    ElseIf IsEmpty Then
                        If InChest Then
                            .AddItem "Place Item"
                            .AddItem "Cancel"
                        Else
                            .AddItem "Cancel"
                        End If
                    Else
                        If InChest Then .AddItem "Catch"
                        If Usable Then .AddItem "Use"
                        .AddItem "Combine"
                        If Not InChest Then .AddItem "Delete"
                        .AddItem "Cancel"
                    End If
                    .Selected(0) = True
                End With

                OptionMode = True
                SelectedRow = 0
        End Select
End Select
End Sub

Private Sub UserForm_Terminate()
    Me.Controls.Clear
End Sub
