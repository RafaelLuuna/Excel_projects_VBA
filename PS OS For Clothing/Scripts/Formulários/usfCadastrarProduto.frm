VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfCadastrarProduto 
   Caption         =   "Cadastro de Produto"
   ClientHeight    =   6390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6315
   OleObjectBlob   =   "usfCadastrarProduto.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfCadastrarProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Linha As Integer
Public Foto As Variant
Private colTamanhos As Collection


Private Sub btConfirmar_Click()
    If Not Checar Then Exit Sub

    
    If Foto = "" Then
        Foto = "Null"
    End If

    
    Dim ColunasTamanhos(5) As Integer
    If CheckBox1 Then
        ColunasTamanhos(0) = 19
        ColunasTamanhos(1) = 20
        ColunasTamanhos(2) = 21
        ColunasTamanhos(3) = 22
        ColunasTamanhos(4) = 23
        ColunasTamanhos(5) = 24
    Else
        ColunasTamanhos(0) = 13
        ColunasTamanhos(1) = 14
        ColunasTamanhos(2) = 15
        ColunasTamanhos(3) = 16
        ColunasTamanhos(4) = 17
        ColunasTamanhos(5) = 18
    End If
    With Planilha3.Range("A:G")
        .Cells(Linha, 1) = Me.combTipoProduto.Value
        .Cells(Linha, 2) = Me.tbDescricao.Value
        .Cells(Linha, 3) = Me.combFornecedor.Value
        .Cells(Linha, 4) = Me.tbEstoqueMinimo.Value
        On Error Resume Next
        If tbEstoqueInicial.Enabled Then
            .Cells(Linha, 5) = Me.tbEstoqueInicial.Value
            .Cells(Linha, ColunasTamanhos(0)) = CDbl(Me.tbTamanho1)
            .Cells(Linha, ColunasTamanhos(1)) = CDbl(Me.tbTamanho2)
            .Cells(Linha, ColunasTamanhos(2)) = CDbl(Me.tbTamanho3)
            .Cells(Linha, ColunasTamanhos(3)) = CDbl(Me.tbTamanho4)
            .Cells(Linha, ColunasTamanhos(4)) = CDbl(Me.tbTamanho5)
            .Cells(Linha, ColunasTamanhos(5)) = CDbl(Me.tbTamanho6)
        End If
        .Cells(Linha, 27) = Foto
        .Cells(Linha, 9) = CDbl(Me.tbValorVenda)
        If .Cells(Linha, 8) = "" Then .Cells(Linha, 8) = 0
        If .Cells(Linha, 10) = "" Then .Cells(Linha, 10) = 0
    End With
    
    MsgBox "Produto cadastrado com sucesso!"
    
    Unload Me
    
End Sub




Private Sub CheckBox1_Click()
    If CheckBox1 Then
        lbTamanho1.Caption = "33-34"
        lbTamanho2.Caption = "34-35"
        lbTamanho3.Caption = "36-37"
        lbTamanho4.Caption = "38-39"
        lbTamanho5.Caption = "40-41"
        lbTamanho6.Caption = "42-43"
    Else
        lbTamanho1.Caption = "PP"
        lbTamanho2.Caption = "P"
        lbTamanho3.Caption = "M"
        lbTamanho4.Caption = "G"
        lbTamanho5.Caption = "GG"
        lbTamanho6.Caption = "GGG"
    End If
End Sub






Private Sub tbTamanho1_Change()
    tbEstoqueInicial = SomarCol(colTamanhos)
End Sub
Private Sub tbTamanho2_Change()
    tbEstoqueInicial = SomarCol(colTamanhos)
End Sub
Private Sub tbTamanho3_Change()
    tbEstoqueInicial = SomarCol(colTamanhos)
End Sub
Private Sub tbTamanho4_Change()
    tbEstoqueInicial = SomarCol(colTamanhos)
End Sub
Private Sub tbTamanho5_Change()
    tbEstoqueInicial = SomarCol(colTamanhos)
End Sub
Private Sub tbTamanho6_Change()
    tbEstoqueInicial = SomarCol(colTamanhos)
End Sub

Private Sub tbTamanho6_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 48 To 57, 96 To 107, 46, 8, 13
        Case Else
            KeyCode = 0
    End Select
End Sub
Private Sub tbTamanho5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 48 To 57, 96 To 107, 46, 8, 13
        Case Else
            KeyCode = 0
    End Select
End Sub
Private Sub tbTamanho4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 48 To 57, 96 To 107, 46, 8, 13
        Case Else
            KeyCode = 0
    End Select
End Sub
Private Sub tbTamanho3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 48 To 57, 96 To 107, 46, 8, 13
        Case Else
            KeyCode = 0
    End Select
End Sub
Private Sub tbTamanho2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 48 To 57, 96 To 107, 46, 8, 13
        Case Else
            KeyCode = 0
    End Select
End Sub
Private Sub tbTamanho1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 48 To 57, 96 To 107, 46, 8, 13
        Case Else
            KeyCode = 0
    End Select
End Sub





Private Sub combTipoProduto_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = 0
End Sub

Private Sub combFornecedor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
KeyCode = 0
End Sub








Private Sub CommandButton4_Click()
    Unload Me
End Sub

Private Sub labelImportarImagem_Click()
    Foto = Application.GetOpenFilename(FileFilter:="Foto(*.jpg), *.jpg")
    If Foto = False Then Exit Sub
    Me.Image1.Picture = LoadPicture(Foto)
End Sub

Private Sub labelCancelar_Click()
    Me.Image1.Picture = LoadPicture("")
    Foto = ""
End Sub

Private Sub tbDescricao_Change()
tbDescricao.Text = UCase(tbDescricao.Text)
End Sub


Private Sub tbEstoqueMinimo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 48 To 57, 96 To 107, 46, 8, 13
        Case Else
            KeyCode = 0
    End Select
End Sub

Private Sub tbEstoqueInicial_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub

Private Sub tbValorVenda_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarValor(tbValorVenda, KeyCode)
End Sub









Private Sub UserForm_Initialize()

    Dim i As Integer
    combTipoProduto.Clear
    combFornecedor.Clear
    For i = 2 To Planilha3.Range("K:K").Cells.SpecialCells(xlCellTypeConstants).Count
        combTipoProduto.AddItem Planilha3.Cells(i, 11).Value
    Next
    For i = 2 To Planilha7.Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count
        combFornecedor.AddItem Planilha7.Cells(i, 1).Value
    Next
    
    Set colTamanhos = New Collection
    colTamanhos.Add tbTamanho1
    colTamanhos.Add tbTamanho2
    colTamanhos.Add tbTamanho3
    colTamanhos.Add tbTamanho4
    colTamanhos.Add tbTamanho5
    colTamanhos.Add tbTamanho6
    
End Sub

Private Sub btNovoTipo_Click()
    Planilha3.Cells(Planilha3.Range("K:K").Cells.SpecialCells(xlCellTypeConstants).Count + 1, 11) = UCase(InputBox("Digite o nome do tipo de produto", "Novo Tipo de Produto"))
    combTipoProduto.Clear
    For i = 2 To Planilha3.Range("K:K").Cells.SpecialCells(xlCellTypeConstants).Count
        combTipoProduto.AddItem Planilha3.Cells(i, 11).Value
    Next
End Sub

Private Sub btNovoFornecedor_Click()
    usfCadastrarFornecedor.Linha = Planilha7.Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
    usfCadastrarFornecedor.Show
    combFornecedor.Clear
    For i = 2 To Planilha7.Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count
        combFornecedor.AddItem Planilha7.Cells(i, 1).Value
    Next
End Sub

Private Function Checar() As Boolean
    



    If tbDescricao = "" Then
        GoTo CheckFalse
    ElseIf combTipoProduto.Text = "" Then
        GoTo CheckFalse
    ElseIf combFornecedor.Text = "" Then
        GoTo CheckFalse
    ElseIf tbEstoqueMinimo.Text = "" Then
        GoTo CheckFalse
    Else
        GoTo CheckTrue
    End If
    
CheckFalse:
        MsgBox "Algum campo não está preenchido", vbCritical
        Checar = False
        Exit Function
        
CheckTrue:
    Checar = True
    
End Function

