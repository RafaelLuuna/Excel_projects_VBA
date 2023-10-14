VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfrmCompraMercadoria 
   Caption         =   "Registrar compra de mercadoria"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9525.001
   OleObjectBlob   =   "usfrmCompraMercadoria.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfrmCompraMercadoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TempLinProduto As Integer

Private Sub CommandButton2_Click()
    usfrmPesquisarProduto.Show
    Call AtualizarProdTemp(Functions.LinProduto)
    Frame1.Visible = True
End Sub

Private Sub CommandButton3_Click()
    If lbProdutoSelecionado = "" Then
        MsgBox "Selecione um produto"
    Else
        Dim iList As Integer
        iList = ListCompra.ListCount
        ListCompra.AddItem ""
        With ListCompra
            .List(iList, 0) = TempLinProduto
            .List(iList, 1) = lbProdutoSelecionado.Caption
            .List(iList, 2) = tbQnt
            .List(iList, 3) = tbValor
            .List(iList, 4) = CDbl(tbValor) * tbQnt
            .List(iList, 5) = tbValorVenda
        End With
        
        TempLinProduto = 0
        
        lbProdutoSelecionado.Caption = ""
        tbQnt = 0
        tbValor = "0,00"
        tbValorVenda = "0,00"
        
        Frame1.Visible = False
        
    End If
End Sub

Private Sub CommandButton4_Click()
    Dim iList As Integer
    For iList = 1 To ListCompra.ListCount - 1
        If ListCompra.Selected(iList) Then
            ListCompra.RemoveItem (iList)
            Debug.Print iList
        End If
    Next
End Sub

Private Sub CommandButton5_Click()
    usfrmCadastroProduto.Show
End Sub

Private Sub AtualizarProdTemp(LinProd)
    TempLinProduto = LinProd
    If LinProd = 0 Then
        lbProdutoSelecionado.Caption = ""
        tbQnt = 0
        tbValor = "0,00"
        tbValorVenda = "0,00"
    Else
        lbProdutoSelecionado.Caption = Planilha3.Cells(LinProd, 2)
        tbQnt = 0
        tbValor = Format(Planilha3.Cells(LinProd, 3), "0.00")
        tbValorVenda = Format(Planilha3.Cells(LinProd, 4), "0.00")
    End If
End Sub

Private Sub CommandButton6_Click()
    TempLinProduto = 0
    
    lbProdutoSelecionado.Caption = ""
    tbQnt = 0
    tbValor = "0,00"
    tbValorVenda = "0,00"
    Frame1.Visible = False
End Sub

Private Sub tbCodigo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Select Case KeyCode
        Case 13
            Dim i As Integer
            For i = 2 To Planilha3.UsedRange.Rows.Count
                If Planilha3.Cells(i, 1) = tbCodigo Then
                    Call AtualizarProdTemp(i)
                    tbCodigo = ""
                    Frame1.Visible = True
                    Exit Sub
                End If
            Next
            Call AtualizarProdTemp(0)
            MsgBox "Produto não encontrado"
        Case Else
    End Select
End Sub

Private Sub tbDataCompra_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarNumeros.FormatarData(tbDataCompra, KeyCode)
End Sub

Private Sub tbQnt_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarNumeros.LimitarTamanho(tbQnt, KeyCode, 10)
End Sub

Private Sub tbValor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarNumeros.FormatarValor(tbValor, KeyCode)
End Sub

Private Sub tbValorVenda_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarNumeros.FormatarValor(tbValorVenda, KeyCode)
End Sub

Private Sub UserForm_Initialize()
    ListCompra.Clear
    ListCompra.AddItem ""
    ListCompra.List(0, 0) = "#"
    ListCompra.List(0, 1) = "Descrição do produto"
    ListCompra.List(0, 2) = "Qnt."
    ListCompra.List(0, 3) = "Preço de custo (uni.)"
    ListCompra.List(0, 4) = "Preço de custo (total)"
    ListCompra.List(0, 5) = "Valor de venda"
End Sub
