VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfCadastrarCliente 
   Caption         =   "Cadastrar Cliente"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6150
   OleObjectBlob   =   "usfCadastrarCliente.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfCadastrarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Linha As Integer

Private Sub btConfirmar_Click()
    Planilha1.Unprotect "1234"
    If Not Checar Then Exit Sub
    
    With Planilha1
        .Cells(Linha, 1) = tbNome.Text
        .Cells(Linha, 2) = tbRua.Text
        .Cells(Linha, 3) = tbNumero.Text
        .Cells(Linha, 4) = tbTelefone.Text
        .Cells(Linha, 5) = tbCPF.Text
        .Cells(Linha, 6) = combSexo.Text
        .Cells(Linha, 7) = Date
    End With
    
    MsgBox "Cliente cadastrado com sucesso"
    
    Planilha1.Protect "1234"
    
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
    Image2.Picture = LoadPicture(Application.ThisWorkbook.Path & "\image source\logo.bmp")
    
    combSexo.AddItem "MASCULINO"
    combSexo.AddItem "FEMININO"
    combSexo.AddItem "NÃO DEFINIDO"
End Sub



Private Sub combSexo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub

Private Sub tbCPF_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarCPF(tbCPF, KeyCode)
End Sub

Private Sub tbNome_Change()
    tbNome = UCase(tbNome.Text)
End Sub


Private Sub tbNumero_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call LimitarTamanho(tbNumero, KeyCode, 5)
End Sub


Private Sub tbRua_Change()
    tbRua = UCase(tbRua.Text)
End Sub


Private Sub tbTelefone_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarTelefone(tbTelefone, KeyCode)
End Sub


Private Sub btCancelar_Click()
    Linha = 0
    Unload Me
End Sub

Private Function Checar() As Boolean
    If tbNome.Text = "" Then
        GoTo CheckFalse
    ElseIf tbRua.Text = "" Then
        GoTo CheckFalse
    ElseIf tbNumero.Text = "" Then
        GoTo CheckFalse
    ElseIf Len(tbCPF.Value) < 14 And Len(tbCPF.Value) > 0 Then
        GoTo CheckFalse
    ElseIf combSexo.Text = "" Then
        GoTo CheckFalse
    ElseIf Len(tbTelefone.Value) < 14 And Len(tbTelefone.Value) > 0 Then
        GoTo CheckFalse
    Else
        GoTo CheckTrue
    End If
    
CheckFalse:
        MsgBox "Algum campo está vazio ou não está preenchido corretamente", vbCritical
        Checar = False
        Exit Function
        
CheckTrue:
    Checar = True
    
End Function

