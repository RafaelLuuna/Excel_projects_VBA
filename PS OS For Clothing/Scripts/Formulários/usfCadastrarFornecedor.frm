VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfCadastrarFornecedor 
   Caption         =   "UserForm1"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6330
   OleObjectBlob   =   "usfCadastrarFornecedor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfCadastrarFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Linha As Integer

Private Sub btCancelar_Click()
    Unload Me
End Sub

Private Sub btConfirmar_Click()
    With Planilha7
        .Cells(Linha, 1) = Me.tbNome
        .Cells(Linha, 2) = Me.tbRua
        .Cells(Linha, 3) = Me.tbNumero
        .Cells(Linha, 4) = Me.tbTelefone1
        .Cells(Linha, 5) = Me.tbTelefone2
        .Cells(Linha, 6) = Me.tbEmail
        .Cells(Linha, 7) = Date
    End With
    MsgBox Me.tbNome & " foi cadastrado com sucesso"
    Unload Me
End Sub

Private Sub tbNome_Change()
    tbNome = UCase(tbNome)
End Sub

Private Sub tbRua_Change()
    tbRua = UCase(tbRua)
End Sub

Private Sub tbNumero_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call LimitarTamanho(tbNumero, KeyCode, 4)
End Sub

Private Sub tbTelefone2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarTelefone(tbTelefone2, KeyCode)
End Sub

Private Sub tbTelefone1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarTelefone(tbTelefone1, KeyCode)
End Sub





Private Sub UserForm_Initialize()
    Image2.Picture = LoadPicture(Application.ThisWorkbook.Path & "\image source\logo.bmp")
End Sub
