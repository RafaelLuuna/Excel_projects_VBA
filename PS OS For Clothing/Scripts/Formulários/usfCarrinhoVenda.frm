VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfCarrinhoVenda 
   Caption         =   "UserForm1"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8340.001
   OleObjectBlob   =   "usfCarrinhoVenda.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfCarrinhoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub btConfirmar_Click()
    Unload Me
End Sub

Private Sub btExcluir_Click()
    Dim i As Integer
    For i = 1 To UBound(lbCarrinho.List)
        If lbCarrinho.Selected(i) Then
            If MsgBox("Deseja excluir " & lbCarrinho.List(i, 0) & " " & lbCarrinho.List(i, 1), vbYesNo) = vbNo Then Exit Sub
            With usfVenda
                .ValorTotal = lbTotal.Caption - lbCarrinho.List(i, 3)
                .Carrinho.Remove (i)
            End With
        End If
    Next
    Call UserForm_Initialize
End Sub





Private Sub lbCarrinho_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer

    
    lbCarrinho.Clear
    
    With Me.lbCarrinho
        .ColumnCount = 7
        .ColumnWidths = "72;125;35;40;40;60;30"
        .ColumnHeads = False
        .AddItem
        .List(0, 0) = "[TIPO]"
        .List(0, 1) = "[DESCRICAO]"
        .List(0, 2) = "[TAM.]"
        .List(0, 3) = "[VALOR]"
        .List(0, 4) = "[DESC.]"
        .List(0, 5) = "[PAGAMENTO]"
        .List(0, 6) = "[QNT.]"
    End With
    
    For i = 1 To usfVenda.Carrinho.Count
        With Me.lbCarrinho
            .AddItem
            .List(i, 0) = usfVenda.Carrinho.Item(i).Tipo
            .List(i, 1) = usfVenda.Carrinho.Item(i).Nome
            .List(i, 2) = usfVenda.Carrinho.Item(i).Tamanho
            .List(i, 3) = Format(usfVenda.Carrinho.Item(i).Valor, "#,##0.00") & " R$"
            .List(i, 4) = Format(usfVenda.Carrinho.Item(i).Desconto, "#,##0.00")
            .List(i, 5) = usfVenda.Carrinho.Item(i).MetodoPagamento
            .List(i, 6) = usfVenda.Carrinho.Item(i).Qnt
        End With
    Next
    
    lbTotal.Caption = Format(usfVenda.ValorTotal, "#,##0.00")

End Sub

Private Sub UserForm_Terminate()
    usfVenda.lbTotalCarrinho = Format(usfVenda.ValorTotal, "#,##0.00")
End Sub


