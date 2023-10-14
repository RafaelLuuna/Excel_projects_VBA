VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfVenda 
   Caption         =   "REALIZAR VENDA"
   ClientHeight    =   8340.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8055
   OleObjectBlob   =   "usfVenda.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Linha As Integer
Public Carrinho As Collection
Public ValorTotal As Double

Private Sub btCadastratCliente_Click()
    Planilha1.Unprotect "1234"
    usfCadastrarCliente.Linha = Planilha1.Range("A:A").Cells.SpecialCells(xlCellTypeConstants).Count + 1
    usfCadastrarCliente.Show
    Planilha1.Unprotect "1234"
    With Planilha1.ListObjects("tabCLIENTES").Range '|LIMPAR FILTROS tabCLIENTES
        .AutoFilter Field:=1                        '|
        .AutoFilter Field:=2                        '|
        .AutoFilter Field:=3                        '|
        .AutoFilter Field:=4                        '|
        .AutoFilter Field:=5                        '|
        .AutoFilter Field:=6                        '|
        .AutoFilter Field:=7                        '|
    End With                                        '|
'----------------------------------------------------|
    combCliente.Clear
    For i = 2 To Planilha1.Range("A:G").End(xlDown).Row
        combCliente.AddItem Planilha1.Cells(i, 1)
    Next
    Planilha1.Protect "1234"
End Sub

Private Sub CheckBox1_Click()
    If CheckBox1 Then
        With combTamanhos
            .Clear
            .AddItem "33-34"
            .AddItem "35-36"
            .AddItem "37-38"
            .AddItem "39-40"
            .AddItem "41-42"
            .AddItem "43-44"
            .ListIndex = 0
        End With
    Else
        With combTamanhos
            .Clear
            .AddItem "PP"
            .AddItem "P"
            .AddItem "M"
            .AddItem "G"
            .AddItem "GG"
            .AddItem "GGG"
            .ListIndex = 0
        End With
    End If
End Sub









Private Sub btAddCarrinho_Click()

    If Not Checar Then Exit Sub
    
    Dim oProduto As New clsProdutos
    With oProduto
        .Nome = tbDescricao.Text
        .Tipo = tbTipo.Text
        .Linha = Planilha3.Range("B:B").Find(oProduto.Nome).Row
        .Qnt = tbQnt.Value
        .Valor = tbPrecoVendido.Value
        .MetodoPagamento = combMetodoPagamento.Text
        .Desconto = CDbl(lbDescontoAcrescimo.Caption)
        .Tamanho = combTamanhos.Text
    End With
    
    If ChecarDulplicidade(Carrinho, oProduto) Then Exit Sub
    
    With usfVenda
        .ValorTotal = .ValorTotal + .tbPrecoVendido.Value
        .Carrinho.Add oProduto
        .lbTotalCarrinho = Format(.ValorTotal, "#,##0.00")
        .tbTipo = ""
        .tbDescricao = ""
        .tbQnt = "1"
        .tbPrecoVendido = "0,00"
        .lbDescontoAcrescimo = "0,00"
        .lbDescontoAcrescimoMargem = "0,0"
        .lbPrecoVenda = "0,00"
        .ListBox1.Clear
        Call Filtrar
    End With
    
    
    
End Sub


Private Sub btConfirmar_Click()
    Dim i, iCol As Integer
    iCol = 1
    For i = Linha To Linha + Carrinho.Count - 1
        With Planilha4
            .Cells(i, 1) = CDate(tbData.Text)
            .Cells(i, 2) = Carrinho.Item(iCol).Tipo
            .Cells(i, 3) = Carrinho.Item(iCol).Nome
            .Cells(i, 4) = Carrinho.Item(iCol).Tamanho
            .Cells(i, 5) = Carrinho.Item(iCol).Qnt
            .Cells(i, 6) = Carrinho.Item(iCol).Valor
            .Cells(i, 7) = Carrinho.Item(iCol).Desconto
            .Cells(i, 8) = usfVenda.combCliente.Text
            .Cells(i, 9) = usfVenda.combMetodoPagamento.Text
        End With
        With Planilha3
            .Cells(Carrinho.Item(iCol).Linha, .Range("M1:X1").Find(Carrinho.Item(iCol).Tamanho).Column) = .Cells(Carrinho.Item(iCol).Linha, .Range("M1:X1").Find(Carrinho.Item(iCol).Tamanho).Column) - Carrinho.Item(iCol).Qnt
        End With
        iCol = iCol + 1
    Next
    Unload Me
End Sub




Private Sub CommandButton4_Click()
    Unload Me
End Sub

Private Sub lbCarrinho_Click()
usfCarrinhoVenda.Show
End Sub


Private Sub ListBox1_Change()
    Dim i As Integer
    For i = 1 To UBound(ListBox1.List)
        If ListBox1.Selected(i) Then
            tbTipo.Text = ListBox1.List(i, 0)
            tbDescricao.Text = ListBox1.List(i, 1)
            lbPrecoVenda.Caption = Format(Planilha3.Cells(Planilha3.Range("B:B").Find(tbDescricao.Value).Row, 9).Value, "#,##0.00")
            If Planilha3.Cells(Planilha3.Range("A:G").Find(ListBox1.List(i, 1)).Row, 27).Value = "Null" Then
                Image3.Picture = LoadPicture("")
            Else
                Image3.Picture = LoadPicture(Planilha3.Cells(Planilha3.Range("A:G").Find(ListBox1.List(i, 1)).Row, 27).Value)
            End If
            tbQnt.Text = 1
            Exit For
        Else
            tbTipo.Text = ""
            tbDescricao.Text = ""
            Image3.Picture = LoadPicture("")
            tbQnt.Text = ""
        End If
    Next
End Sub









Private Sub combCliente_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub


Private Sub tbData_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarData(tbData, KeyCode)
End Sub

Private Sub tbDescricao_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub

Private Sub tbPrecoVendido_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarValor(tbPrecoVendido, KeyCode)
    Call CalcularLabels
End Sub

Private Sub CalcularLabels()
    On Error Resume Next
    lbPrecoVenda.Caption = Format(Planilha3.Cells(Planilha3.Range("B:B").Find(tbDescricao.Text).Row, 9).Value * tbQnt.Value, "#,##0.00")
    lbDescontoAcrescimo.Caption = Format(tbPrecoVendido.Value - CDbl(lbPrecoVenda.Caption), "#,##0.00")
    lbDescontoAcrescimoMargem.Caption = Format(CDbl(lbDescontoAcrescimo.Caption) / CDbl(lbPrecoVenda) * 100, "#,##0.00")

    If CDbl(lbDescontoAcrescimo.Caption) < Planilha3.Cells(Planilha3.Range("B:B").Find(tbDescricao.Text).Row, 10).Value * tbQnt.Value * -1 Then
        lbDescontoAcrescimo.ForeColor = &H80&
        lbDescontoAcrescimoMargem.ForeColor = &H80&
        rs.ForeColor = &H80&
        rsM.ForeColor = &H80&
    Else
        lbDescontoAcrescimo.ForeColor = &H8000&
        lbDescontoAcrescimoMargem.ForeColor = &H8000&
        rs.ForeColor = &H8000&
        rsM.ForeColor = &H8000&
    End If
    
End Sub

Private Sub tbQnt_Change()
    Call CalcularLabels
End Sub

Private Sub combTamanhos_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub

Private Sub tbTipo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub


Private Sub combFiltroTipo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub

Private Sub combFornecedores_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub

Private Sub combFiltroTipo_Change()
    Call Filtrar
End Sub

Private Sub combFornecedores_Change()
    Call Filtrar
End Sub


Private Sub tbPesquisa_Change()
    tbPesquisa = UCase(tbPesquisa)
    Call Filtrar
End Sub



Private Sub UserForm_Initialize()
    Image2.Picture = LoadPicture(Application.ThisWorkbook.Path & "\image source\logo.bmp")
    
    Set Carrinho = New Collection
    tbData = Date
    
    With Planilha3.ListObjects("tabESTOQUE").Range  '|LIMPAR FILTROS tabESTOQUE
        .AutoFilter Field:=1                        '|
        .AutoFilter Field:=2                        '|
        .AutoFilter Field:=3                        '|
        .AutoFilter Field:=4                        '|
        .AutoFilter Field:=5                        '|
        .AutoFilter Field:=6                        '|
        .AutoFilter Field:=7                        '|
    End With                                        '|
'----------------------------------------------------|
    

    
    Dim i As Integer
    
    combFiltroTipo.AddItem "*[TODOS]*"
    combFornecedores.AddItem "*[TODOS]*"
    
    For i = 2 To Planilha3.Range("K:K").End(xlDown).Row
        combFiltroTipo.AddItem Planilha3.Cells(i, 11)
    Next
    For i = 2 To Planilha7.Range("A:A").End(xlDown).Row
        combFornecedores.AddItem Planilha7.Cells(i, 1)
    Next
    
    With combTamanhos
        .AddItem "PP"
        .AddItem "P"
        .AddItem "M"
        .AddItem "G"
        .AddItem "GG"
        .AddItem "GGG"
    End With
        
    combTamanhos.ListIndex = 0
    combFiltroTipo.ListIndex = 0
    combFornecedores.ListIndex = 0
    
    Call Filtrar
    
    With Planilha1.ListObjects("tabCLIENTES").Range '|LIMPAR FILTROS tabCLIENTES
        .AutoFilter Field:=1                        '|
        .AutoFilter Field:=2                        '|
        .AutoFilter Field:=3                        '|
        .AutoFilter Field:=4                        '|
        .AutoFilter Field:=5                        '|
        .AutoFilter Field:=6                        '|
        .AutoFilter Field:=7                        '|
    End With                                        '|
'----------------------------------------------------|

    For i = 2 To Planilha1.Range("A:G").End(xlDown).Row
        combCliente.AddItem Planilha1.Cells(i, 1)
    Next
    
    
    With combMetodoPagamento
        .AddItem "DINHEIRO"
        .AddItem "DEBITO"
        .AddItem "CREDITO"
    End With
    
End Sub

Private Sub Filtrar()
    Dim i, iList As Integer
    
    Dim FilterIndex(2) As Boolean
    FilterIndex(0) = False
    FilterIndex(1) = False
    FilterIndex(2) = False
    If Not combFiltroTipo.Text = "*[TODOS]*" Then FilterIndex(0) = True
    If Not combFornecedores.Text = "*[TODOS]*" Then FilterIndex(1) = True
    If Not tbPesquisa.Text = "" Then FilterIndex(2) = True

    ListBox1.Clear
    tbTipo.Text = ""
    tbDescricao.Text = ""
    Image3.Picture = LoadPicture("")
    tbQnt.Text = ""
    
    
    With ListBox1                       '| PREPARAR CABEÇALHO
        .AddItem                        '|
        .List(0, 0) = "[TIPO]"          '|
        .List(0, 1) = "[DESCRIÇÃO]"     '|
        .List(0, 2) = "[FORNECEDOR]"    '|
    End With                            '|
    iList = 1                           '|
'----------------------------------------|
    

    
    For i = 2 To Planilha3.Range("A:G").End(xlDown).Row
FiltroNome:
        If Not FilterIndex(2) Then GoTo FiltroTipo
        
        
        If InStr(1, Planilha3.Range("A:G").Cells(i, 2).Value, tbPesquisa.Text) > 0 Then
            
            If Not FilterIndex(0) And Not FilterIndex(1) Then
                GoTo GetList
            Else: GoTo FiltroTipo
            End If
            
        Else: GoTo NextItem
        End If
        
        
        
FiltroTipo:
        If Not FilterIndex(0) Then GoTo FiltroFornecedor

        If Planilha3.Range("A:G").Cells(i, 1) = combFiltroTipo.Text Then

            If FilterIndex(1) Then
                GoTo FiltroFornecedor
            Else: GoTo GetList
            End If

        Else: GoTo NextItem
        End If


FiltroFornecedor:
        If Not FilterIndex(1) Then GoTo GetList

        If Planilha3.Range("A:G").Cells(i, 3) = combFornecedores.Text Then
            GoTo GetList
        Else: GoTo NextItem
        End If


GetList:
        With Me.ListBox1
            .AddItem
            .List(iList, 0) = Planilha3.Cells(i, 1)
            .List(iList, 1) = Planilha3.Cells(i, 2)
            .List(iList, 2) = Planilha3.Cells(i, 3)
        End With
        iList = iList + 1
NextItem:
        Next i
    


End Sub

Private Function ChecarDulplicidade(Col As Collection, oProduto As clsProdutos) As Boolean
    Dim i As Integer
    For i = 1 To Col.Count
        If Col.Item(i).Nome = oProduto.Nome And Col.Item(i).Tamanho = oProduto.Tamanho And Col.Item(i).MetodoPagamento = oProduto.MetodoPagamento Then
                MsgBox oProduto.Nome & "(" & oProduto.Tamanho & ") já está no carrinho", vbInformation
                ChecarDulplicidade = True
                Exit Function
        End If
    Next
End Function

Private Function Checar() As Boolean
    
    Dim i, iColumn As Integer
    i = Planilha3.Range("B:B").Find(Me.tbDescricao.Text).Row
    iColumn = Planilha3.Range("M1:X1").Find(Me.combTamanhos.Text).Column
    If Planilha3.Cells(i, iColumn) = "" Then Planilha3.Cells(i, iColumn) = 0
    If Me.tbQnt.Text = "" Then Me.tbQnt = 0
    If Planilha3.Cells(i, iColumn) - Me.tbQnt < 0 Then GoTo CheckFalseEstoque:




    If tbTipo.Text = "" Then
        GoTo CheckFalse
    ElseIf tbDescricao.Text = "" Then
        GoTo CheckFalse
    ElseIf tbQnt.Text = "" Then
        GoTo CheckFalse
    ElseIf tbPrecoVendido.Text = "0,00" Then
        GoTo CheckFalse
    ElseIf Len(tbData) < 10 Then
        GoTo CheckFalse
    ElseIf combTamanhos.Text = "" Then
        GoTo CheckFalse
    ElseIf combMetodoPagamento.Text = "" Then
        GoTo CheckFalse
    ElseIf combCliente.Text = "" Then
        GoTo CheckFalse
    Else
        GoTo CheckTrue
    End If
    

CheckFalse:
                    MsgBox "Algum campo não está preenchido", vbCritical
                    Checar = False
                    Exit Function


CheckFalseEstoque:
                    MsgBox "Estoque Insuficiente", vbCritical
                    MsgBox "Restam: " & Planilha3.Cells(i, iColumn) & " " & Me.tbTipo & " " & Me.tbDescricao & " tamanho '" & Me.combTamanhos.Text & "'"
                    Checar = False
                    Exit Function


CheckTrue:
                    Checar = True
                    Exit Function
                

                    
                    
                    
    
End Function
