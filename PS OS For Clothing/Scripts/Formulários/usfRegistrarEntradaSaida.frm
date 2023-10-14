VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfRegistrarEntradaSaida 
   Caption         =   "UserForm1"
   ClientHeight    =   1815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8115
   OleObjectBlob   =   "usfRegistrarEntradaSaida.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfRegistrarEntradaSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit:
Public Coluna As Integer

Private Sub btConfirmar_Click()
    If Not Checar Then Exit Sub
    Call Planilha6.Registrar(Coluna, tbValor.Value, tbDescricao.Value, tbData.Value)
    Unload Me
End Sub


Private Sub tbData_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarData(tbData, KeyCode)
End Sub

Private Sub tbValor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarValor(tbValor, KeyCode)
End Sub

Private Sub UserForm_Initialize()
    tbData.Text = Date
    tbValor.SetFocus
End Sub

Private Function Checar() As Boolean
    
    If Len(tbData) < 10 Then
        GoTo CheckFalse
    ElseIf tbValor = "0,00" Then
        GoTo CheckFalse
    ElseIf tbDescricao = "" Then
        GoTo CheckFalse
    Else
        GoTo CheckTrue
    End If
    
    
CheckFalse:
    MsgBox "Algum campo não foi preenchido corretamente", vbInformation
    Checar = False
    Exit Function
    
CheckTrue:
    Checar = True
    
End Function
