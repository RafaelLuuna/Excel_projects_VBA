VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfrmTalk 
   Caption         =   "Name"
   ClientHeight    =   1470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8310.001
   OleObjectBlob   =   "usfrmTalk.frx":0000
End
Attribute VB_Name = "usfrmTalk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iTalk As Integer

Private Sub UserForm_Initialize()
    Me.Width = GameScreen.Width
    Me.Top = GameScreen.Top + GameScreen.Height - Me.Height
    Me.Left = GameScreen.Left
    iTalk = 1
    ScriptFunctions.RenderScript (iTalk)
End Sub


Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim ActionID
    Select Case KeyCode
        Case 70, 90, 13, 32
            If iTalk = WorksheetFunction.VLookup(DATA.ActualScript & "," & iTalk, ScriptData.Range("C:D"), 2, False) Then
                Unload Me
                Exit Sub
            End If
            iTalk = iTalk + 1
            ScriptFunctions.RenderScript (iTalk)
            ScriptFunctions.DoAction (iTalk)
        Case Else
    End Select
End Sub
