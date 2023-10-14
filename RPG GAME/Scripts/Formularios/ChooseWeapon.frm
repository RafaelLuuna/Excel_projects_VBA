VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChooseWeapon 
   Caption         =   "ChooseWeapon"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3720
   OleObjectBlob   =   "ChooseWeapon.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChooseWeapon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private AxesArray() As Weapon
Private AxeCount As Integer

Private Sub CommandButton3_Click()
    If AxeCount = 0 Then
        MsgBox "You have no Axes to use"
        Exit Sub
    End If
    
    With AxesArray(SpinButton1.Value)
        Call Wood_Game.SetWeapon(.Dmg, .Weight, .Precision, .Durabillity, .NumSlot)
        Wood_Game.lbChances = .Durabillity
    End With
    Unload Me
    Wood_Game.Show
    
End Sub

Private Sub SpinButton1_Change()
If SpinButton1.Value > AxeCount - 1 Then
    SpinButton1 = AxeCount - 1
    Exit Sub
End If
lbWpName = AxesArray(SpinButton1.Value).ID
lbPrecision = AxesArray(SpinButton1.Value).Precision
lbDmg = AxesArray(SpinButton1.Value).Dmg
lbWeight = AxesArray(SpinButton1.Value).Weight
lbDurabillity = AxesArray(SpinButton1.Value).Durabillity

On Error Resume Next
imgWeapon.Picture = LoadPicture(Application.ThisWorkbook.Path & "\texture\item\" & AxesArray(SpinButton1.Value).ID & ".gif")

End Sub

Private Sub SpinButton1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Call CommandButton3_Click
End Sub

Private Sub UserForm_Initialize()
Dim i, iWp As Integer

For i = 2 To WpData.Range("A:A").SpecialCells(xlCellTypeConstants).Count
    If InventoryFunctions.CountItem(1, WpData.Cells(i, 1)) > 0 Then
        For iWp = 1 To InventoryFunctions.CountItem(1, WpData.Cells(i, 1))
            AxeCount = AxeCount + 1
        Next
    End If
Next

ReDim AxesArray(AxeCount)
AxeCount = 0

For i = 2 To WpData.Range("A:A").SpecialCells(xlCellTypeConstants).Count
    If InventoryFunctions.CountItem(1, WpData.Cells(i, 1)) > 0 Then
        For iWp = 1 To InventoryFunctions.CountItem(1, WpData.Cells(i, 1))
            With AxesArray(AxeCount)
                .NumID = iWp
                .ID = WpData.Cells(i, 1)
                If iWp = 1 Then
                    .NumSlot = InventoryFunctions.FindItem(1, .ID)
                Else
                    .NumSlot = InventoryFunctions.FindItem(1, .ID, AxesArray(AxeCount - 1).NumSlot + 1)
                End If
                .Dmg = WpData.Cells(i, 2)
                .Weight = WpData.Cells(i, 3)
                .Precision = WpData.Cells(i, 4)
                .Durabillity = DATA.InventoryArray(1).InventorySlots(.NumSlot).Durabillity
            End With
            AxeCount = AxeCount + 1
        Next
    End If
Next

lbWpName = AxesArray(0).ID
lbPrecision = AxesArray(0).Precision
lbDmg = AxesArray(0).Dmg
lbWeight = AxesArray(0).Weight
lbDurabillity = AxesArray(0).Durabillity

SpinButton1.SetFocus

On Error Resume Next
imgWeapon.Picture = LoadPicture(Application.ThisWorkbook.Path & "\texture\item\" & AxesArray(0).ID & ".gif")


End Sub

