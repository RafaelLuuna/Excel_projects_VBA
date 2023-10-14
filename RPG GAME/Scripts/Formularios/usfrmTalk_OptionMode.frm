VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfrmTalk_OptionMode 
   ClientHeight    =   1920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4845
   OleObjectBlob   =   "usfrmTalk_OptionMode.frx":0000
End
Attribute VB_Name = "usfrmTalk_OptionMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    Dim i As Integer
    Dim objOption As Control
    
    DATA.Talk_SelectedOption = 1

    Me.Height = (DATA.Talk_OptionNum * 35) + 29.25

    For i = 1 To DATA.Talk_OptionNum
        Set objOption = Me.Controls.Add("Forms.Label.1", "Option " & i)
        With objOption
            .Height = 30
            .Width = Me.InsideWidth - 8
            .Left = 4
            .Top = (i - 1) * 35
            .BackColor = &H8000000F
            If i = DATA.Talk_SelectedOption Then
                .BorderStyle = 1
                .BackColor = &HE0E0E0
            End If
            .Caption = DATA.Talk_OptionTexts.Item(i)
        End With
    Next

    Me.Left = usfrmTalk.Left + usfrmTalk.Width - Me.Width
    Me.Top = usfrmTalk.Top - Me.Height
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 87, 38
            If Not DATA.Talk_SelectedOption = 1 Then
                'Remove a seleção atual
                Me.Controls("Option " & DATA.Talk_SelectedOption).BackColor = &H8000000F
                Me.Controls("Option " & DATA.Talk_SelectedOption).BorderStyle = 0
                DATA.Talk_SelectedOption = DATA.Talk_SelectedOption - 1
                'Atualiza a seleção atual
                Me.Controls("Option " & DATA.Talk_SelectedOption).BackColor = &HE0E0E0
                Me.Controls("Option " & DATA.Talk_SelectedOption).BorderStyle = 1
            End If
        Case 83, 40
            If Not DATA.Talk_SelectedOption = DATA.Talk_OptionNum Then
                'Remove a seleção atual
                Me.Controls("Option " & DATA.Talk_SelectedOption).BackColor = &H8000000F
                Me.Controls("Option " & DATA.Talk_SelectedOption).BorderStyle = 0
                DATA.Talk_SelectedOption = DATA.Talk_SelectedOption + 1
                'Atualiza a seleção atual
                Me.Controls("Option " & DATA.Talk_SelectedOption).BackColor = &HE0E0E0
                Me.Controls("Option " & DATA.Talk_SelectedOption).BorderStyle = 1
            End If
        Case 70, 90, 13, 32
            Unload Me
            If usfrmTalk.iTalk = WorksheetFunction.VLookup(DATA.ActualScript & "," & usfrmTalk.iTalk, ScriptData.Range("C:E"), 3, False) Then
                Unload usfrmTalk
                Exit Sub
            End If
            usfrmTalk.iTalk = usfrmTalk.iTalk + 1
            ScriptFunctions.RenderScript (usfrmTalk.iTalk)
            ScriptFunctions.DoAction (usfrmTalk.iTalk)
        Case Else
    End Select
End Sub
