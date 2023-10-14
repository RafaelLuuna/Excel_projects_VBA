VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Wood_Game 
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9255.001
   OleObjectBlob   =   "Wood_Game.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Wood_Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LoopState As Boolean
Private HitState As Boolean

Private Point1 As Integer
Private Point2 As Integer
Private Velocity As Double
Private MaxVelocity As Double
Private Distancia(1) As Double
Private ObjectPosition As Double

Private Goal As Integer

Private AttackPoints As Integer
Private Weight As Integer
Private Chances As Integer
Private Margem As Integer

Public AxeSlot As Integer

Private Sub UserForm_Initialize()
    Call Update_Tree_Life

    LoopState = False
    HitState = False
End Sub

Public Sub SetWeapon(Dmg As Integer, Wgt As Integer, Precision As Integer, Durability As Integer, Optional WpSlot As Integer = 9999)
    AttackPoints = Dmg
    Weight = Wgt
    Margem = 50 - Precision
    Chances = Durability
    AxeSlot = WpSlot
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Select Case KeyCode
    Case 32
        
        lbTextStart.Visible = False
        
        
        
        If HitState Then
            HitState = False
            
            lbGoal.Visible = False
            
            Dim HitPosition As Integer
            Dim Precisao As Double
            
            HitPosition = ObjectPosition
            Precisao = HitPosition - Goal
            
            If Precisao < 0 Then Precisao = -Precisao
            If Precisao > Margem Then
                Precisao = 0
            Else
                Precisao = (Margem - Precisao) / Margem
            End If
            lbPrecisao = Format(Precisao, "#,##0.00%")
            
            DATA.Tree_Life = DATA.Tree_Life - (AttackPoints * Precisao)
            If DATA.Tree_Life < 0 Then DATA.Tree_Life = 0
            lbDmg = Format(AttackPoints * Precisao, "#,##0.00")
            Call Update_Tree_Life
            
            ObjectPosition = imgAxe.Left
            MaxVelocity = 20
            Point2 = 257
            
            LoopState = True
            Call MoveObject(imgAxe)
            
            If DATA.Tree_Life <= 0 Then
                MsgBox "Parabéns!!"
                PlayerControls.Win = True
                LoopState = False
                Unload Me
            End If
            
            If Chances <= 0 Then
                Chances = 0
                lbChances = Chances
                LoopState = False
                If InventoryFunctions.CountWeapons = 0 Then
                    MsgBox "You have no more axes to use."
                    MsgBox "End Game"
                Else
                    Unload Me
                    ChooseWeapon.Show
                End If
                Unload Me
                Exit Sub
            End If
        Else
            If LoopState Then Exit Sub
            Chances = Chances - 1
            If Chances < 0 Then
                Chances = 0
                lbChances = Chances
                LoopState = False
                MsgBox "Suas chances acabaram!"
                Unload Me
                Exit Sub
            End If
            
            If Not AxeSlot = 0 Then Call InventoryFunctions.ChangeSlot(1, AxeSlot, , , Chances)
            
            HitState = True
            lbChances = Chances
            
            lbGoal.Visible = True
            Goal = CInt(Rnd() * 140) + 40
            lbGoal.Width = Margem + 40
            lbGoal.Left = Goal - (lbGoal.Width / 2)
            
            ObjectPosition = imgAxe.Left
            MaxVelocity = (20 - Weight) / 5
            Point2 = 10
            
            LoopState = True
            Call MoveObject(imgAxe)
        End If

    Case 13
        If LoopState Then
                LoopState = False
        Else:   LoopState = True
        End If
    Case 38
        Weight = Weight + 1
        lbDebug = "MaxVelocity = " & 4 + Weight
    Case 40
        MaxVelocityBonus = MaxVelocityBonus - 1
        lbDebug = "MaxVelocity = " & 4 + Weight
    
    Case 39
        Margem = Margem + 1
        lbDebug = "Margem de erro = " & Margem
    Case 37
        Margem = Margem - 1
        lbDebug = "Margem de erro = " & Margem
    Case Else
        MsgBox KeyCode
End Select
End Sub

Private Sub Update_Tree_Life()
    On Error Resume Next
    lbTree_Life.Width = (1 - ((DATA.Max_Tree_life - DATA.Tree_Life) / DATA.Max_Tree_life)) * 312
    lbTree_Life = DATA.Tree_Life
    If DATA.Tree_Life = 0 Then lbTree_Life.Width = 0
End Sub

Private Sub MoveObject(ByRef Object)
        On Error Resume Next
        Point1 = ObjectPosition
        Distancia(0) = (Point1 - Point2) / 2
        If Distancia(0) < 0 Then Distancia(0) = Distancia(0) * -1
        
        Do
        DoEvents
            Distancia(1) = imgAxe.Left - Point2
            If Distancia(1) < 0 Then
                Distancia(1) = Distancia(1) + 1
                Velocity = (-(Distancia(1) + Distancia(0)) ^ 2 * (MaxVelocity / (Distancia(0) ^ 2))) + MaxVelocity
                ObjectPosition = ObjectPosition + Velocity
            ElseIf Distancia(1) > 0 Then
                Distancia(1) = Distancia(1) - 1
                Velocity = ((Distancia(1) - Distancia(0)) ^ 2 * (MaxVelocity / (Distancia(0) ^ 2))) - MaxVelocity
                ObjectPosition = ObjectPosition + Velocity
            End If
            
            imgAxe.Left = ObjectPosition
            
            If imgAxe.Left > 400 Or ObjectPosition > 400 Or imgAxe.Left < -400 Or ObjectPosition < -400 Then
                ObjectPosition = 252
                imgAxe.Left = 252
                LoopState = False
            End If
            
            If ObjectPosition < Point2 + 1.1 And ObjectPosition > Point2 - 1.1 Then
                lbTextStart.Visible = True
                LoopState = False
            End If
        lbDebug = "[LoopState = " & LoopState & "] | [ObjectPosition = " & CInt(ObjectPosition) & "] | [Goal = " & Goal & "]"
        
        Call TimeOut(0.0001)
        Loop Until LoopState = False

End Sub

Private Sub TimeOut(Duration As Double)
Dim StartTime As Double
StartTime = Timer
Do
DoEvents
Loop Until (Timer - StartTime) > Duration
End Sub
