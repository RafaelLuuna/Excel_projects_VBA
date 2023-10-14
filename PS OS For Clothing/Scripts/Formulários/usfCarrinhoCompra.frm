VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfCarrinhoCompra 
   Caption         =   "UserForm1"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8340.001
   OleObjectBlob   =   "usfCarrinhoCompra.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfCarrinhoCompra"
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
            With usfCOMPRA
                .ValorTotal = lbTotal.Caption - lbCarrinho.List(i, 3)
                .Carrinho.Remove (i)
            End With
        End If
    Next
    Call UserForm_Initialize
End Sub


Private Sub UserForm_Initialize()
    Dim i As Integer

    
    lbCarrinho.Clear
    
    With Me.lbCarrinho
        .ColumnCount = 5
        .ColumnWidths = "72;180;60;60;30"
        .ColumnHeads = False
        .AddItem
        .List(0, 0) = "[TIPO]"
        .List(0, 1) = "[DESCRICAO]"
        .List(0, 2) = "[TAMANHO]"
        .List(0, 3) = "[VALOR]"
        .List(0, 4) = "[QNT.]"
    End With
    
    For i = 1 To usfCOMPRA.Carrinho.Count
        With Me.lbCarrinho
            .AddItem
            .List(i, 0) = usfCOMPRA.Carrinho.Item(i).Tipo
            .List(i, 1) = usfCOMPRA.Carrinho.Item(i).Nome
            .List(i, 2) = usfCOMPRA.Carrinho.Item(i).Tamanho
            .List(i, 3) = Format(usfCOMPRA.Carrinho.Item(i).Valor, "#,##0.00") & " R$"
            .List(i, 4) = usfCOMPRA.Carrinho.Item(i).Qnt
        End With
    Next
    
    lbTotal.Caption = Format(usfCOMPRA.ValorTotal, "#,##0.00")

End Sub

Private Sub UserForm_Terminate()
    usfCOMPRA.lbTotalCarrinho = Format(usfCOMPRA.ValorTotal, "#,##0.00")
End Sub


