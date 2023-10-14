Attribute VB_Name = "PlayerControls"
Public Moving As Boolean
Public TextureBehindPlayer As String

Public Win As Boolean


Public Sub Interact()
Dim ID As String
    With DATA.ActualScene
    ID = .Layer2.Cells(PlayerVar.Direction.Y + .YPOS - 1, PlayerVar.Direction.x + .XPOS - 1)
    End With

Select Case ID
    Case "Closed_Chest"
        Call Load.ChangeID(PlayerVar.Direction.x + DATA.ActualScene.XPOS - 1, PlayerVar.Direction.Y + DATA.ActualScene.YPOS - 1, 2, "Opened_Chest")
        Call Load.InventorySlots(WorksheetFunction.Match(ChestNames.Find(PlayerVar.Direction.x + DATA.ActualScene.XPOS - 1, PlayerVar.Direction.Y + DATA.ActualScene.YPOS - 1, ActualScene.Layer2.Name), InventoryData.Range("1:1"), 0), True, "Chest")
        usfrmInventory.Show
        Call Load.ChangeID(PlayerVar.Direction.x + DATA.ActualScene.XPOS - 1, PlayerVar.Direction.Y + DATA.ActualScene.YPOS - 1, 2, "Closed_Chest")
    Case "Opened_Chest"
        Call Load.ChangeID(PlayerVar.Direction.x + DATA.ActualScene.XPOS - 1, PlayerVar.Direction.Y + DATA.ActualScene.YPOS - 1, 2, "Closed_Chest")
    Case "Trunk"
        If InventoryFunctions.CountWeapons = 0 Then
            MsgBox "Is a beautiful tree, if i had an axe maybe i can get some wood..."
        Else
            If InventoryFunctions.CountItem(1, "Wood") = 0 And InventoryFunctions.FindItem(1, "Null") = 0 Then
                MsgBox "I'm afraid there's no space in my bag to pick up these woods."
            Else
                Win = False
                DATA.Tree_Life = 100
                DATA.Max_Tree_life = DATA.Tree_Life
                ChooseWeapon.Show
                If Win = True Then
                    If InventoryFunctions.CountItem(1, "Wood") = 0 Then
                        Call InventoryFunctions.ChangeSlot(1, InventoryFunctions.FindItem(1, "Null"), "Wood", 3, 0)
                    Else
                        Call InventoryFunctions.ChangeSlot(1, InventoryFunctions.FindItem(1, "Wood"), , DATA.InventoryArray(1).InventorySlots(InventoryFunctions.FindItem(1, "Wood")).Qnt + 3)
                    End If
                    '////Remove sprites da árvore
                    Call Load.ChangeID(DATA.PlayerVar.Direction.x + DATA.ActualScene.XPOS - 1, DATA.PlayerVar.Direction.Y - 1 + DATA.ActualScene.YPOS - 1, 3, "Air")
                    Call Load.ChangeID(DATA.PlayerVar.Direction.x + DATA.ActualScene.XPOS - 1, DATA.PlayerVar.Direction.Y - 2 + DATA.ActualScene.YPOS - 1, 3, "Air")
                    Call Load.ChangeID(DATA.PlayerVar.Direction.x + DATA.ActualScene.XPOS - 1, DATA.PlayerVar.Direction.Y - 3 + DATA.ActualScene.YPOS - 1, 3, "Air")
                    Call Load.ChangeID(DATA.PlayerVar.Direction.x + 1 + DATA.ActualScene.XPOS - 1, DATA.PlayerVar.Direction.Y - 1 + DATA.ActualScene.YPOS - 1, 3, "Air")
                    Call Load.ChangeID(DATA.PlayerVar.Direction.x - 1 + DATA.ActualScene.XPOS - 1, DATA.PlayerVar.Direction.Y - 1 + DATA.ActualScene.YPOS - 1, 3, "Air")
                    Call Load.ChangeID(DATA.PlayerVar.Direction.x + 1 + DATA.ActualScene.XPOS - 1, DATA.PlayerVar.Direction.Y - 2 + DATA.ActualScene.YPOS - 1, 3, "Air")
                    Call Load.ChangeID(DATA.PlayerVar.Direction.x - 1 + DATA.ActualScene.XPOS - 1, DATA.PlayerVar.Direction.Y - 2 + DATA.ActualScene.YPOS - 1, 3, "Air")
                    Call Load.ChangeID(DATA.PlayerVar.Direction.x + DATA.ActualScene.XPOS - 1, DATA.PlayerVar.Direction.Y + DATA.ActualScene.YPOS - 1, 2, "Cut_Trunk")
                End If
            End If
        End If
    Case "CraftTable"
        Crafting.Show
    Case "NPC_1"
        Dim ScriptSequence As Collection
        Set ScriptSequence = New Collection
        ScriptSequence.Add 1
        
        DATA.ActualScript = ScriptFunctions.NextScriptID(ScriptSequence)
        
        DATA.ScriptCheck.Add DATA.ActualScript
        'MsgBox Data.ActualScript.ScriptID
        usfrmTalk.Show
End Select
End Sub


Private Function DetectCollision() As Boolean
    Dim TargetID As String
    With DATA.ActualScene
    On Error Resume Next
    TargetID = .Layer2.Cells(PlayerVar.Direction.Y + .YPOS - 1, PlayerVar.Direction.x + .XPOS - 1)
    End With
    If Left(TargetID, 4) = "Door" Then TargetID = "Door"
    
    If Err.Number = 1004 Then
        DetectCollision = True
        Exit Function
    End If
    
    Select Case TargetID
        Case "Door", "Air", ""
            DetectCollision = False
        Case Else
            DetectCollision = True
    End Select
End Function





Public Sub MoveRight()
    If Moving Then Exit Sub
    Moving = True
    
    Call PlayerVar.SetDirection(PlayerVar.Position.x + 1, PlayerVar.Position.Y)
    If PlayerVar.Direction.x > DATA.xArraySize Then Call PlayerVar.SetDirection(DATA.xArraySize, PlayerVar.Position.Y)
    Call Render.Player
    
    If DetectCollision Then
        Moving = False
        Exit Sub
    End If
    
    Do
    DoEvents
        GameScreen.Controls("Player").Left = GameScreen.Controls("Player").Left + 1
        If GameScreen.Controls("Player").Left > DATA.SpriteArray(PlayerVar.Direction.x, PlayerVar.Direction.Y, 1).XPOS Then
            GameScreen.Controls.Item("Player").Left = DATA.SpriteArray(PlayerVar.Direction.x, PlayerVar.Direction.Y, 1).XPOS + 0.5
        End If
    Call TimeOut(0.01)
    Loop Until GameScreen.Controls("Player").Left >= DATA.SpriteArray(PlayerVar.Direction.x, PlayerVar.Direction.Y, 1).XPOS
    
    DATA.SpriteArray(PlayerVar.Position.x, PlayerVar.Position.Y, 2).ID = "Air"
    
    If Not PlayerVar.Position.x + 1 > DATA.xArraySize Then Call PlayerVar.SetPosition(PlayerVar.Position.x + 1, PlayerVar.Position.Y)
    Call PlayerVar.SetDirection(PlayerVar.Position.x + 1, PlayerVar.Position.Y)
    If PlayerVar.Direction.x > DATA.xArraySize + 1 Then Call PlayerVar.SetDirection(DATA.xArraySize + 1, PlayerVar.Position.Y)
    
    TextureBehindPlayer = DATA.SpriteArray(PlayerVar.Position.x, PlayerVar.Position.Y, 1).ID
    DATA.SpriteArray(PlayerVar.Position.x, PlayerVar.Position.Y, 2).ID = "Player"
    
    Moving = False

    Call GameScreen.AfterWalk

End Sub

Public Sub MoveLeft()
    If Moving Then Exit Sub
    Moving = True
    
    Call PlayerVar.SetDirection(PlayerVar.Position.x - 1, PlayerVar.Position.Y)
    'If PlayerVar.Direction.x < 1 Then Call PlayerVar.SetDirection(1, PlayerVar.Position.y)
    Call Render.Player
    
    If DetectCollision Then
        Moving = False
        Exit Sub
    End If
    
    Do
    DoEvents
        GameScreen.Controls("Player").Left = GameScreen.Controls("Player").Left - 1
        If GameScreen.Controls("Player").Left < DATA.SpriteArray(PlayerVar.Direction.x, PlayerVar.Direction.Y, 1).XPOS Then
            GameScreen.Controls.Item("Player").Left = DATA.SpriteArray(PlayerVar.Direction.x, PlayerVar.Direction.Y, 1).XPOS - 0.5
        End If
    Call TimeOut(0.01)
    Loop Until GameScreen.Controls("Player").Left <= DATA.SpriteArray(PlayerVar.Direction.x, PlayerVar.Direction.Y, 1).XPOS
    
    DATA.SpriteArray(PlayerVar.Position.x, PlayerVar.Position.Y, 2).ID = "Air"
    
    Call PlayerVar.SetPosition(PlayerVar.Position.x - 1, PlayerVar.Position.Y)
    Call PlayerVar.SetDirection(PlayerVar.Position.x - 1, PlayerVar.Position.Y)
    If PlayerVar.Direction.x < 1 Then Call PlayerVar.SetDirection(0, PlayerVar.Position.Y)
    
    TextureBehindPlayer = DATA.SpriteArray(PlayerVar.Position.x, PlayerVar.Position.Y, 1).ID
    DATA.SpriteArray(PlayerVar.Position.x, PlayerVar.Position.Y, 2).ID = "Player"
    
    Moving = False

    Call GameScreen.AfterWalk

End Sub

Public Sub MoveUp()
    If Moving Then Exit Sub
    Moving = True
    
    Call PlayerVar.SetDirection(PlayerVar.Position.x, PlayerVar.Position.Y - 1)
    'If PlayerVar.Direction.y < 1 Then Call PlayerVar.SetDirection(PlayerVar.Position.x, 1)
    Call Render.Player
    
    If DetectCollision Then
        Moving = False
        Exit Sub
    End If
    
    Do
    DoEvents
        GameScreen.Controls("Player").Top = GameScreen.Controls("Player").Top - 1
        If GameScreen.Controls("Player").Top < DATA.SpriteArray(PlayerVar.Direction.x, PlayerVar.Direction.Y, 1).YPOS Then
            GameScreen.Controls.Item("Player").Top = DATA.SpriteArray(PlayerVar.Direction.x, PlayerVar.Direction.Y, 1).YPOS - 0.5
        End If
    Call TimeOut(0.01)
    Loop Until GameScreen.Controls("Player").Top <= DATA.SpriteArray(PlayerVar.Direction.x, PlayerVar.Direction.Y, 1).YPOS
    
    DATA.SpriteArray(PlayerVar.Position.x, PlayerVar.Position.Y, 2).ID = "Air"
    
    Call PlayerVar.SetPosition(PlayerVar.Position.x, PlayerVar.Position.Y - 1)
    Call PlayerVar.SetDirection(PlayerVar.Position.x, PlayerVar.Position.Y - 1)
    If PlayerVar.Direction.Y < 1 Then Call PlayerVar.SetDirection(PlayerVar.Position.x, 0)
    
    TextureBehindPlayer = DATA.SpriteArray(PlayerVar.Position.x, PlayerVar.Position.Y, 1).ID
    DATA.SpriteArray(PlayerVar.Position.x, PlayerVar.Position.Y, 2).ID = "Player"
    
    Moving = False

    Call GameScreen.AfterWalk

End Sub

Public Sub MoveDown()
    If Moving Then Exit Sub
    Moving = True
    
    Call PlayerVar.SetDirection(PlayerVar.Position.x, PlayerVar.Position.Y + 1)
    If PlayerVar.Direction.Y > DATA.yArraySize Then Call PlayerVar.SetDirection(PlayerVar.Position.x, DATA.yArraySize)
    Call Render.Player
    
    If DetectCollision Then
        Moving = False
        Exit Sub
    End If
    
    Do
    DoEvents
        GameScreen.Controls("Player").Top = GameScreen.Controls("Player").Top + 1
        If GameScreen.Controls("Player").Top > DATA.SpriteArray(PlayerVar.Direction.x, PlayerVar.Direction.Y, 1).YPOS Then
            GameScreen.Controls.Item("Player").Top = DATA.SpriteArray(PlayerVar.Direction.x, PlayerVar.Direction.Y, 1).YPOS + 0.5
        End If
    Call TimeOut(0.01)
    Loop Until GameScreen.Controls("Player").Top >= DATA.SpriteArray(PlayerVar.Direction.x, PlayerVar.Direction.Y, 1).YPOS
    
    DATA.SpriteArray(PlayerVar.Position.x, PlayerVar.Position.Y, 2).ID = "Air"
    
    If Not PlayerVar.Position.Y + 1 > DATA.yArraySize Then Call PlayerVar.SetPosition(PlayerVar.Position.x, PlayerVar.Position.Y + 1)
    Call PlayerVar.SetDirection(PlayerVar.Position.x, PlayerVar.Position.Y + 1)
    If PlayerVar.Direction.Y > DATA.yArraySize + 1 Then Call PlayerVar.SetDirection(PlayerVar.Position.x, DATA.yArraySize + 1)
    
    
    TextureBehindPlayer = DATA.SpriteArray(PlayerVar.Position.x, PlayerVar.Position.Y, 1).ID
    DATA.SpriteArray(PlayerVar.Position.x, PlayerVar.Position.Y, 2).ID = "Player"
    
    Moving = False
    
    Call GameScreen.AfterWalk
    
End Sub

Private Sub TimeOut(Duration As Double)
Dim StartTime As Double
StartTime = Timer
Do
DoEvents
Loop Until (Timer - StartTime) > Duration
End Sub



