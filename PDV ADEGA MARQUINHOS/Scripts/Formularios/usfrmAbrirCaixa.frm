VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfrmAbrirCaixa 
   Caption         =   "Abertura do caixa"
   ClientHeight    =   1905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4110
   OleObjectBlob   =   "usfrmAbrirCaixa.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfrmAbrirCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If TextBox1 = "" Then
        MsgBox "Digite o nome do responsável"
        Exit Sub
    End If
    
    Planilha5.Range("B1") = TextBox1.Value
    Planilha5.Range("B6") = CDbl(TextBox2.Value)
    Planilha5.Range("B4") = Now
    Planilha5.Range("C1") = ""
    
    ThisWorkbook.Save
    
    Unload Me
End Sub

Private Sub TextBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarNumeros.FormatarValor(TextBox2, KeyCode)
End Sub

Private Sub UserForm_Initialize()
    Planilha5.Range("C1") = "Cancelar abertura do caixa"
End Sub
