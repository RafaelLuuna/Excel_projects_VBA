VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfrmQnt 
   Caption         =   " "
   ClientHeight    =   1635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2385
   OleObjectBlob   =   "usfrmQnt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfrmQnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MaxValue As Integer

Private Sub CommandButton1_Click()
If Not TextBox1.Value = 1 Then TextBox1.Value = TextBox1.Value - 1
End Sub

Private Sub CommandButton2_Click()
TextBox1.Value = TextBox1.Value + 1
End Sub

Private Sub CommandButton3_Click()
Unload Me
End Sub

Private Sub CommandButton4_Click()
TextBox1.Value = MaxValue
End Sub

Private Sub TextBox1_Change()
If TextBox1 = "" Then TextBox1 = 1
If TextBox1.Value > MaxValue Then TextBox1 = MaxValue
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
If KeyCode = 13 Then Unload Me
Select Case KeyCode
    Case 87, 38
        TextBox1 = MaxValue
    Case 65, 37
        If Not TextBox1.Value = 1 Then TextBox1.Value = TextBox1.Value - 1
    Case 68, 39
        TextBox1.Value = TextBox1.Value + 1
    Case 70, 13, 90
        Unload Me
End Select
End Sub

Private Sub UserForm_Initialize()
DATA.VarQnt = 0
End Sub

Private Sub UserForm_Terminate()
DATA.VarQnt = CInt(TextBox1.Value)
End Sub
