VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GameMenu 
   Caption         =   "UserForm1"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7605
   OleObjectBlob   =   "GameMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GameMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_Wood_Game_Click()
Call Wood_Game.SetWeapon(50, 15, 10, 30, 0)
Wood_Game.Show
End Sub

Private Sub CommandButton1_Click()
GameScreen.Show
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub UserForm_Click()
    Dim wbName As Window

    Set wbName = ActiveWorkbook.Windows(1) 'You can use Windows("[Workbook Name]") as well

    wbName.Visible = True

End Sub
