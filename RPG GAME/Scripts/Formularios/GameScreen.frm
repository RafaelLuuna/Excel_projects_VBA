VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GameScreen 
   Caption         =   "UserForm1"
   ClientHeight    =   8175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "GameScreen.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GameScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LastWorldPosition As Position
Private iScene As Position
Private PlayerPos As Position

Private Sub UserForm_Initialize()
    LastWorldPosition.x = 8
    LastWorldPosition.Y = 8
    Call Load.InventoryObjects
    Set DATA.ScriptCheck = New Collection
    
    With DATA.ActualScene
        Set .Layer1 = Mundo1
        Set .Layer2 = Mundo2
        Set .Layer3 = Mundo3
        .XPOS = 1
        .YPOS = 1
    End With
    PlayerPos.x = 8
    PlayerPos.Y = 10
    Me.Caption = "World"
    Call LoadWorld
End Sub

Private Sub LoadWorld()
    
    Me.Controls.Clear
    Call Load.Display
    Call Load.Player(PlayerPos.x, PlayerPos.Y)
    
    Call Load.DisplayLayers
    Call Load.Scene(DATA.ActualScene.Layer1, DATA.ActualScene.Layer2, DATA.ActualScene.Layer3, DATA.ActualScene.XPOS, DATA.ActualScene.YPOS)
    'Call Load.DEBUG_LB
    Call Load.HUD
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

Select Case KeyCode
    Case 87, 38
        PlayerControls.MoveUp
    Case 65, 37
        PlayerControls.MoveLeft
    Case 83, 40
        PlayerControls.MoveDown
    Case 68, 39
        PlayerControls.MoveRight
    Case 69, 67
        Call Load.InventorySlots(1, False, "Inventory")
        usfrmInventory.Show
    Case 70, 90, 13
        Call PlayerControls.Interact
End Select

'GameScreen.Controls.Item("DEBUG_LB").Caption = PlayerVar.Position.x & "//" & PlayerVar.Direction.x & "|   |" & PlayerVar.Position.y & "//" & PlayerVar.Direction.y
'GameScreen.Controls.Item("DEBUG_LB").Caption = Data.SpriteArray(PlayerVar.Direction.X, PlayerVar.Direction.Y, 2).ID
'GameScreen.Controls.Item("DEBUG_LB").Caption = PlayerControls.TextureBehindPlayer
End Sub

Public Sub AfterWalk()
If PlayerVar.Position.x = xArraySize Then
    DATA.ActualScene.XPOS = DATA.ActualScene.XPOS + (xArraySize - 1)
    PlayerPos.x = 1
    PlayerPos.Y = PlayerVar.Position.Y
    Call LoadWorld
End If
If PlayerVar.Position.x = 0 Then
    DATA.ActualScene.XPOS = DATA.ActualScene.XPOS - (xArraySize - 1)
    PlayerPos.x = xArraySize - 1
    PlayerPos.Y = PlayerVar.Position.Y
    Call LoadWorld
End If
If PlayerVar.Position.Y = yArraySize Then
    DATA.ActualScene.YPOS = DATA.ActualScene.YPOS + (yArraySize - 1)
    PlayerPos.x = PlayerVar.Position.x
    PlayerPos.Y = 1
    Call LoadWorld
End If
If PlayerVar.Position.Y = 0 Then
    DATA.ActualScene.YPOS = DATA.ActualScene.YPOS - (yArraySize - 1)
    PlayerPos.x = PlayerVar.Position.x
    PlayerPos.Y = yArraySize - 1
    Call LoadWorld
End If

Select Case PlayerControls.TextureBehindPlayer
    Case "Door"
        LastWorldPosition.x = PlayerVar.Position.x
        LastWorldPosition.Y = PlayerVar.Position.Y + 1
        With DATA.ActualScene
            Set .Layer1 = Casa_1
            Set .Layer2 = Casa_2
            Set .Layer3 = Casa_3
            .XPOS = 1
            .YPOS = 1
        End With
        PlayerPos.x = 10
        PlayerPos.Y = 10
        Me.Caption = "House"
        Call LoadWorld
    Case "LightedWoodenFloor"
        With DATA.ActualScene
            Set .Layer1 = Mundo1
            Set .Layer2 = Mundo2
            Set .Layer3 = Mundo3
            .XPOS = 1
            .YPOS = 1
        End With
        PlayerPos.x = 6
        PlayerPos.Y = 8
        Me.Caption = "World"
        Call LoadWorld
End Select
End Sub


