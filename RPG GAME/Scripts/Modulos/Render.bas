Attribute VB_Name = "Render"
Public Sub UpdateHud()
    GameScreen.Controls("Wallet_HUD").Caption = Format(WalletData.Cells(2, 2), "$ 0.00")
End Sub

Public Sub Backgroung()
    Dim ix, iy As Integer
    Dim Texture As String
    For iy = 1 To yArraySize
    For ix = 1 To xArraySize
        Texture = TextureTest(DATA.SpriteArray(ix, iy, 1).ID, "Block", ".jpg")
        GameScreen.Controls.Item("Sprite" & ix & "," & iy & "," & 1).Picture = LoadPicture(Application.ThisWorkbook.Path & "\texture\block\" & Texture & ".jpg")
    Next
    Next
End Sub
Public Sub Layers()
    Dim ix, iy, iz As Integer
    Dim Texture As String
    Dim TexturePath As String
    For iz = 2 To zArraySize
    For iy = 1 To yArraySize
    For ix = 1 To xArraySize
        Texture = TextureTest(DATA.SpriteArray(ix, iy, iz).ID, "Block", ".gif")
        TexturePath = Application.ThisWorkbook.Path & "\texture\block\" & Texture & ".gif"
        If Texture = "Null" Then
            Texture = TextureTest(DATA.SpriteArray(ix, iy, iz).ID, "Entity", ".gif")
            TexturePath = Application.ThisWorkbook.Path & "\texture\entity\" & Texture & ".gif"
        End If
        GameScreen.Controls.Item("Sprite" & ix & "," & iy & "," & iz).Picture = LoadPicture(TexturePath)
    Next
    Next
    Next
End Sub

Public Sub Sprite(xCoord, yCoord, zCoord)
    Dim Texture As String
    
    Select Case zCoord
        Case 1
            Texture = DATA.ActualScene.Layer1.Cells(yCoord + DATA.ActualScene.YPOS - 1, xCoord + DATA.ActualScene.XPOS - 1)
            GameScreen.Controls.Item("Sprite" & xCoord & "," & yCoord & "," & 1).Picture = LoadPicture(Application.ThisWorkbook.Path & "\texture\block\" & Texture & ".jpg")
        Case 2
            Texture = DATA.ActualScene.Layer2.Cells(yCoord + DATA.ActualScene.YPOS - 1, xCoord + DATA.ActualScene.XPOS - 1)
            GameScreen.Controls.Item("Sprite" & xCoord & "," & yCoord & "," & 2).Picture = LoadPicture(Application.ThisWorkbook.Path & "\texture\block\" & Texture & ".gif")
        Case 3
            Texture = DATA.ActualScene.Layer3.Cells(yCoord + DATA.ActualScene.YPOS - 1, xCoord + DATA.ActualScene.XPOS - 1)
            GameScreen.Controls.Item("Sprite" & xCoord & "," & yCoord & "," & 3).Picture = LoadPicture(Application.ThisWorkbook.Path & "\texture\block\" & Texture & ".gif")
    End Select
End Sub


Public Sub Player()
    Select Case PlayerVar.Direction.x
        Case Is > PlayerVar.Position.x
            DATA.PlayerDirection = "Right"
        Case Is < PlayerVar.Position.x
            DATA.PlayerDirection = "Left"
        Case Else
            Select Case PlayerVar.Direction.Y
                Case Is > PlayerVar.Position.Y
                    DATA.PlayerDirection = "Front"
                Case Is < PlayerVar.Position.Y
                    DATA.PlayerDirection = "Back"
            End Select
    End Select
    GameScreen.Controls.Item("Player").Picture = LoadPicture(Application.ThisWorkbook.Path & "\texture\entity\Player_" & DATA.PlayerDirection & ".gif")
End Sub

Public Sub InventorySlots(InventoryID As Integer)
    Dim i As Integer
    For i = 1 To DATA.InventoryArray(InventoryID).InventorySize
        usfrmInventory.Controls.Item("Slot" & i).Picture = LoadPicture(Application.ThisWorkbook.Path & "\texture\item\" & DATA.InventoryArray(InventoryID).InventorySlots(i).ID & ".gif")
        usfrmInventory.Controls.Item("Slot_Qnt" & i).Caption = CStr(DATA.InventoryArray(InventoryID).InventorySlots(i).Qnt)
    Next
End Sub


Private Function TextureTest(Texture As String, TextureType As String, Format As String) As String
On Error Resume Next
Dim oControl As Control
Set oControl = GameScreen.Controls.Add("Forms.Image.1", "TextureTest", False)
GameScreen.Controls.Item("TextureTest").Picture = LoadPicture(Application.ThisWorkbook.Path & "\texture\" & TextureType & "\" & Texture & Format)
TextureTest = Texture
If Err.Number = 53 Then TextureTest = "Null"
GameScreen.Controls.Remove ("TextureTest")
End Function
