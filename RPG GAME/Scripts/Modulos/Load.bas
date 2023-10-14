Attribute VB_Name = "Load"
Private xSize As Double
Private ySize As Double
Private nTrail As Integer

Public Sub HUD()
    Dim oControl As Control
    Set oControl = GameScreen.Controls.Add("Forms.Label.1", "Wallet_HUD")
    With oControl
        .Top = 0
        .Left = 0
        .Height = 25
        .Width = 90
        .Font.Bold = True
        .Font.Size = 14
        .ForeColor = &HFFFF&
        .BackStyle = 0
    End With
    Call Render.UpdateHud
End Sub

Public Sub Scene(Layer1 As Worksheet, Layer2 As Worksheet, Layer3 As Worksheet, XPOS As Integer, YPOS As Integer)
    For x = 1 To xArraySize
        For Y = 1 To yArraySize
            DATA.SpriteArray(x, Y, 1).ID = Layer1.Cells(YPOS + Y - 1, XPOS + x - 1)
            DATA.SpriteArray(x, Y, 2).ID = Layer2.Cells(YPOS + Y - 1, XPOS + x - 1)
            DATA.SpriteArray(x, Y, 3).ID = Layer3.Cells(YPOS + Y - 1, XPOS + x - 1)
        Next
    Next
    With DATA.ActualScene
        Set .Layer1 = Layer1
        Set .Layer2 = Layer2
        Set .Layer3 = Layer3
        .XPOS = XPOS
        .YPOS = YPOS
    End With
    Render.Backgroung
    Render.Layers
End Sub

Public Sub DisplayLayers()
    HouseCount = 0
    Dim ix, iy, iz As Integer
    Dim oControl As Control
    xSize = GameScreen.InsideWidth / (xArraySize - 1)
    ySize = GameScreen.InsideHeight / (yArraySize - 1)
    For iz = 2 To zArraySize
    For iy = 0 To yArraySize
    For ix = 0 To xArraySize
        On Error Resume Next
        DATA.SpriteArray(ix, iy, iz).ID = "Air"
        DATA.SpriteArray(ix, iy, iz).XPOS = DATA.SpriteArray(ix - 1, iy, iz).XPOS + xSize
        DATA.SpriteArray(ix, iy, iz).YPOS = DATA.SpriteArray(ix, iy - 1, iz).YPOS + ySize
        If iy = 1 Then DATA.SpriteArray(ix, iy, iz).YPOS = 0
        If ix = 1 Then DATA.SpriteArray(ix, iy, iz).XPOS = 0
        If iy = 0 Then DATA.SpriteArray(ix, iy, iz).YPOS = -ySize
        If ix = 0 Then DATA.SpriteArray(ix, iy, iz).XPOS = -xSize
        DATA.SpriteArray(ix, iy, iz).xCoord = ix
        DATA.SpriteArray(ix, iy, iz).yCoord = iy
        Set oControl = GameScreen.Controls.Add("Forms.Image.1", "Sprite" & ix & "," & iy & "," & iz)
        With oControl
            .Top = DATA.SpriteArray(ix, iy, iz).YPOS
            .Left = DATA.SpriteArray(ix, iy, iz).XPOS
            .Height = ySize
            .Width = xSize
            .BorderStyle = fmBorderStyleNone
            .PictureSizeMode = fmPictureSizeModeStretch
            .BackStyle = fmBackStyleTransparent
        End With
    Next
    Next
    Next
    
    Call Render.Layers
    
End Sub
Public Sub Display()
    HouseCount = 0
    Dim ix, iy As Integer
    Dim oControl As Control
    xSize = GameScreen.InsideWidth / (xArraySize - 1)
    ySize = GameScreen.InsideHeight / (yArraySize - 1)
    For iy = 0 To yArraySize
    For ix = 0 To xArraySize
        On Error Resume Next
        DATA.SpriteArray(ix, iy, 1).ID = ""
        DATA.SpriteArray(ix, iy, 1).XPOS = DATA.SpriteArray(ix - 1, iy, 1).XPOS + xSize
        DATA.SpriteArray(ix, iy, 1).YPOS = DATA.SpriteArray(ix, iy - 1, 1).YPOS + ySize
        If iy = 1 Then DATA.SpriteArray(ix, iy, 1).YPOS = 0
        If ix = 1 Then DATA.SpriteArray(ix, iy, 1).XPOS = 0
        If iy = 0 Then DATA.SpriteArray(ix, iy, 1).YPOS = -xSize
        If ix = 0 Then DATA.SpriteArray(ix, iy, 1).XPOS = -ySize
        DATA.SpriteArray(ix, iy, 1).xCoord = ix
        DATA.SpriteArray(ix, iy, 1).yCoord = iy
        Set oControl = GameScreen.Controls.Add("Forms.Image.1", "Sprite" & ix & "," & iy & "," & 1)
        With oControl
            .Top = DATA.SpriteArray(ix, iy, 1).YPOS
            .Left = DATA.SpriteArray(ix, iy, 1).XPOS
            .Height = ySize
            .Width = xSize
            .BorderStyle = fmBorderStyleNone
            .PictureSizeMode = fmPictureSizeModeStretch
        End With
    Next
    Next
    
End Sub

Public Sub Background(TextureID As String)
    Dim ix, iy As Integer
    For iy = 1 To yArraySize
    For ix = 1 To xArraySize
        DATA.SpriteArray(ix, iy).ID = TextureID
    Next
    Next
    Render.Backgroung
End Sub

Public Sub Player(xCoord, yCoord)
    Dim oControl As Control
    Set PlayerVar = New clsPlayer
    Set oControl = GameScreen.Controls.Add("Forms.Image.1", "Player")
    With oControl
        .Top = DATA.SpriteArray(xCoord, yCoord, 1).YPOS
        .Left = DATA.SpriteArray(xCoord, yCoord, 1).XPOS
        .Height = ySize
        .Width = xSize
        .BorderStyle = fmBorderStyleNone
        .PictureSizeMode = fmPictureSizeModeStretch
        .BackStyle = fmBackStyleTransparent
    End With
    PlayerControls.TextureBehindPlayer = DATA.SpriteArray(xCoord, yCoord, 1).ID
    DATA.SpriteArray(xCoord, yCoord, 2).ID = "Player"
    Call PlayerVar.SetPosition(xCoord, yCoord)
    Call PlayerVar.SetDirection(xCoord, yCoord + 1)
    Call Render.Player
End Sub

Public Sub DEBUG_LB()
    Dim oControl As Control
    Set oControl = GameScreen.Controls.Add("Forms.Label.1", "DEBUG_LB")
    With oControl
        .Top = 0
        .Left = 0
        .Height = 20
        .Width = GameScreen.InsideWidth
    End With
End Sub

Public Sub InsertSprite(xCoord, yCoord, TextureID As String, SpriteName As String)
    Dim oControl As Control
    Set oControl = GameScreen.Controls.Add("Forms.Image.1", SpriteName)
    With oControl
        .Height = ySize
        .Width = xSize
        .Top = DATA.SpriteArray(xCoord, yCoord).YPOS
        .Left = DATA.SpriteArray(xCoord, yCoord).XPOS
        .BorderStyle = fmBorderStyleNone
        .PictureSizeMode = fmPictureSizeModeStretch
        .BackStyle = fmBackStyleTransparent
    End With
    DATA.SpriteArray(xCoord, yCoord).ID = SpriteName
    Call Render.SpriteTransparent(xCoord, yCoord, TextureID, SpriteName)
End Sub

Public Sub ChangeID(xCoord, yCoord, zCoord, TextureID As String)
    Select Case zCoord
        Case 1
            DATA.ActualScene.Layer1.Cells(yCoord, xCoord) = TextureID
        Case 2
            DATA.ActualScene.Layer2.Cells(yCoord, xCoord) = TextureID
        Case 3
            DATA.ActualScene.Layer3.Cells(yCoord, xCoord) = TextureID
    End Select
    Call Render.Sprite(xCoord - DATA.ActualScene.XPOS + 1, yCoord - DATA.ActualScene.YPOS + 1, zCoord)
End Sub

Public Sub InventoryObjects()
    Dim oInventory As Inventory
    Dim oItem As Item
    Dim i, iRow As Integer
    
    For i = 1 To InventoryData.Range("1:1").Cells.SpecialCells(xlCellTypeConstants).Count
        With oInventory
            .InventoryName = InventoryData.Cells(1, i)
            .InventorySize = InventoryData.Cells(2, i)
            .ColumnID = i
        End With
        DATA.InventoryArray(i) = oInventory
        For iRow = 3 To DATA.InventoryArray(i).InventorySize + 3
            If InventoryData.Cells(iRow, i) = "" Then
                With oItem
                    .ID = "Null"
                    .Qnt = 0
                    .Durabillity = 0
                End With
            Else
                Dim Spacer1, Spacer2
                With oItem
                    Spacer1 = InStr(1, InventoryData.Cells(iRow, i), ",")
                    .ID = Left(InventoryData.Cells(iRow, i), Spacer1 - 1)
                    Spacer2 = InStr(Spacer1 + 1, InventoryData.Cells(iRow, i), ",")
                    .Qnt = CInt(Mid(InventoryData.Cells(iRow, i), Spacer1 + 1, Spacer2 - Spacer1 - 1))
                    .Durabillity = CInt(Right(InventoryData.Cells(iRow, i), Len(InventoryData.Cells(iRow, i)) - Spacer2))
                End With
            End If
            DATA.InventoryArray(i).InventorySlots(iRow - 2) = oItem
        Next
    Next
End Sub

Public Sub InventorySlots(InventoryID As Integer, Optional InChest = "Omitido", Optional InventoryCaption As String = "Omitido")
    Dim oControl As Control
    Dim Inventory_xSize, Inventory_ySize As Double
    Dim i, iRow, iColumn As Integer
    
    usfrmInventory.Controls.Clear
    
    Inventory_xSize = usfrmInventory.InsideWidth / InventoryDisplaySize
    Inventory_ySize = Inventory_xSize
    
    iRow = 0
    iColumn = 0
    
    x = usfrmInventory.Height - usfrmInventory.InsideHeight
    usfrmInventory.Height = (CInt((DATA.InventoryArray(InventoryID).InventorySize / InventoryDisplaySize) + 3) * Inventory_ySize) - x
    
    For i = 1 To DATA.InventoryArray(InventoryID).InventorySize
    
        Set oControl = usfrmInventory.Controls.Add("Forms.Image.1", "Slot" & i)

        With oControl
            .Height = Inventory_ySize
            .Width = Inventory_xSize
            .Top = Inventory_ySize * iRow
            .Left = Inventory_xSize * iColumn
            .PictureSizeMode = fmPictureSizeModeStretch
            .BackColor = &HFFFFFF
        End With
        
        Set oControl = usfrmInventory.Controls.Add("Forms.Label.1", "Slot_Qnt" & i)
        
        With oControl
            .Height = Inventory_ySize
            .Width = Inventory_xSize
            .Top = Inventory_ySize * iRow
            .Left = Inventory_xSize * iColumn
            .BackStyle = fmBackStyleTransparent
        End With
        
        iColumn = iColumn + 1
        If iColumn = InventoryDisplaySize Then
            iRow = iRow + 1
            iColumn = 0
        End If
        
    Next
    
    usfrmInventory.Controls.Item("Slot" & usfrmInventory.SelectedSlot).BackColor = &HE0E0E0
    Call Render.InventorySlots(InventoryID)
    
    
    usfrmInventory.InventoryID = InventoryID
    If InChest = "Omitido" Then
            usfrmInventory.InChest = usfrmInventory.InChest
    Else:   usfrmInventory.InChest = InChest
    End If
    If InventoryCaption = "Omitido" Then
            usfrmInventory.Caption = usfrmInventory.Caption
    Else:   usfrmInventory.Caption = InventoryCaption
    End If
    
End Sub
