VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usfCompra 
   Caption         =   "UserForm1"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235.001
   OleObjectBlob   =   "usfCompra.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usfCOMPRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Linha As Integer
Public ValorTotal As Double
Public Carrinho As Collection


Private Sub btConfirmar_Click()
    Dim i, iCol As Integer
    iCol = 1
    For i = Linha To Linha + Carrinho.Count - 1
        With Planilha5
            .Cells(i, 1) = CDate(tbData.Text)
            .Cells(i, 2) = Carrinho.Item(iCol).Tipo
            .Cells(i, 3) = Carrinho.Item(iCol).Nome
            .Cells(i, 4) = Carrinho.Item(iCol).Tamanho
            .Cells(i, 5) = Carrinho.Item(iCol).Qnt
            .Cells(i, 6) = Carrinho.Item(iCol).ValorUnitario
            .Cells(i, 7) = Carrinho.Item(iCol).Valor
        End With
        With Planilha3
            .Cells(Carrinho.Item(iCol).Linha, 8) = Carrinho.Item(iCol).ValorUnitario
            .Cells(Carrinho.Item(iCol).Linha, 9) = Carrinho.Item(iCol).PrecoVenda
'            .Cells(Carrinho.Item(iCol).Linha, 10) = Carrinho.Item(iCol).Desconto
            .Cells(Carrinho.Item(iCol).Linha, .Range("M1:X1").Find(Carrinho.Item(iCol).Tamanho).Column) = .Cells(Carrinho.Item(iCol).Linha, .Range("M1:X1").Find(Carrinho.Item(iCol).Tamanho).Column) + Carrinho.Item(iCol).Qnt
        End With
        iCol = iCol + 1
    Next
    Unload Me
End Sub







Private Sub btAddCarrinho_Click()
    If Not Checar Then Exit Sub
    
    Dim oProduto As New clsProdutos
    With oProduto
        .Nome = Me.tbDescricao.Text
        .Tipo = Me.tbTipo.Text
        .Linha = Planilha3.Range("B:B").Find(oProduto.Nome).Row
        .Qnt = Me.tbQnt.Value
        .Valor = CDbl(Me.lbValorTotal.Caption)
        .PrecoVenda = Me.tbPrecoVenda.Value
        .Desconto = Me.tbDescontoPermitido
        .Tamanho = CStr(Me.combTamanhos.Text)
    End With
    
    If ChecarDulplicidade(Carrinho, oProduto) Then Exit Sub
    
    With usfCOMPRA
        .ValorTotal = .ValorTotal + CDbl(lbValorTotal.Caption)
        .lbTotalCarrinho = Format(.ValorTotal, "#,##0.00")
        .tbTipo = ""
        .tbDescricao = ""
        .tbQnt = "1"
        .tbPrecoCompra = "0,00"
        .tbPrecoVenda = "0,00"
        .tbDescontoPermitido = "0,00"
        .lbLucroEmDinheiro.Caption = "0,00"
        .ListBox1.Clear
        Call Filtrar
    End With
    
    Dim i As Integer
    For i = 1 To Carrinho.Count
        If Carrinho.Item(i).Nome = oProduto.Nome And Not Carrinho.Item(i).Tamanho = oProduto.Tamanho Then
            Carrinho.Item(i).PrecoVenda = oProduto.PrecoVenda
        End If
    Next
    
    
    Planilha3.Cells(oProduto.Linha, 10) = oProduto.Desconto
    usfCOMPRA.Carrinho.Add oProduto
    
    MsgBox oProduto.Tipo & " " & oProduto.Nome & " foi adicionada ao carrinho"
    
End Sub

Private Sub btCancelar_Click()
    Unload Me
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

Private Sub lbCarrinho_Click()
    usfCarrinhoCompra.Show
End Sub








Private Sub ListBox1_Change()
    Dim i As Integer
    For i = 1 To UBound(ListBox1.List)
        If ListBox1.Selected(i) Then
            tbTipo.Text = ListBox1.List(i, 0)
            tbDescricao.Text = ListBox1.List(i, 1)
            tbPrecoCompra = Format(ListBox1.List(i, 3), "#,##0.00")
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
            tbPrecoCompra = "0,00"
            Image3.Picture = LoadPicture("")
            tbQnt.Text = ""
        End If
    Next
End Sub

















Private Sub tbPrecoVenda_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarValor(tbPrecoVenda, KeyCode)
    Call CalcularLabels
End Sub

Private Sub tbPrecoCompra_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarValor(tbPrecoCompra, KeyCode)
    Call CalcularLabels
End Sub


Private Sub tbDescontoPermitido_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarValor(tbDescontoPermitido, KeyCode)
    Call CalcularLabels
End Sub

Private Sub tbQnt_Change()
    Call CalcularLabels
End Sub

Private Sub CalcularLabels()
        lbValorTotal.Caption = "0,00"
        lbLucro.Caption = "0,0"
        lbLucroEmDinheiro.Caption = "0,00"
        lbDescontoEmDinheiro.Caption = "0,00"
        lbDescontoPorItem.Caption = "0,0"
    On Error Resume Next
    lbValorTotal.Caption = Format(tbPrecoCompra * tbQnt, "#,##0.00")
    lbLucro.Caption = Format(tbPrecoVenda.Value / tbPrecoCompra * 100 - 100, "#,#0.0")
    lbLucroEmDinheiro.Caption = Format((tbPrecoVenda.Value * tbQnt.Value) - CDbl(lbValorTotal.Caption), "#,##0.00")
    If Not tbDescontoPermitido = "0,00" Then lbDescontoEmDinheiro.Caption = Format(CDbl(lbLucroEmDinheiro.Caption) - (tbDescontoPermitido.Value * tbQnt.Value), "#,##0.00")
    lbDescontoPorItem.Caption = Format(tbDescontoPermitido.Value / tbPrecoVenda.Value * 100, "#,#0.0")

    If tbQnt.Text = "" Then
        
        lbValorTotal.Caption = "0,00"
        lbLucro.Caption = "0,0"
        lbLucroEmDinheiro.Caption = "0,00"
        lbDescontoEmDinheiro.Caption = "0,00"
        lbDescontoPorItem.Caption = "0,0"
    End If
End Sub









Private Sub tbData_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call FormatarData(tbData, KeyCode)
End Sub

Private Sub combTamanhos_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub

Private Sub combCliente_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub

Private Sub tbDescricao_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
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
    Set Carrinho = New Collection
    Image2.Picture = LoadPicture(Application.ThisWorkbook.Path & "\image source\logo.bmp")
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
    tbPrecoCompra = "0,00"
    Image3.Picture = LoadPicture("")
    tbQnt.Text = ""
    
    
    With ListBox1                        '| PREPARAR CABEÇALHO
        .AddItem                         '|
        .List(0, 0) = "[TIPO]"           '|
        .List(0, 1) = "[DESCRIÇÃO]"      '|
        .List(0, 2) = "[FORNECEDOR]"     '|
        .List(0, 3) = "[CUSTO ANTERIOR]" '|
    End With                             '|
    iList = 1                            '|
'-----------------------------------------|
    

    
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
            .List(iList, 3) = Format(Planilha3.Cells(i, 8).Value, "#,##0.00")
        End With
        iList = iList + 1
NextItem:
        Next i
    


End Sub

Private Function ChecarDulplicidade(Col As Collection, oProduto As clsProdutos) As Boolean
    Dim i As Integer
    For i = 1 To Col.Count
        If Col.Item(i).Nome = oProduto.Nome And Col.Item(i).Tamanho = oProduto.Tamanho Then
                MsgBox oProduto.Nome & "(" & oProduto.Tamanho & ") já está no carrinho", vbInformation
                ChecarDulplicidade = True
                Exit Function
        End If
    Next
End Function

Private Function Checar() As Boolean
    If tbTipo.Text = "" Then
        GoTo CheckFalse
    ElseIf tbDescricao.Text = "" Then
        GoTo CheckFalse
    ElseIf tbQnt.Text = "" Then
        GoTo CheckFalse
    ElseIf tbPrecoCompra.Text = "0,00" Then
        GoTo CheckFalse
    ElseIf tbPrecoVenda.Text = "0,00" Then
        GoTo CheckFalse
    ElseIf tbDescontoPermitido.Text = "" Then
        GoTo CheckFalse
    ElseIf Len(tbData) < 10 Then
        GoTo CheckFalse
    ElseIf combTamanhos.Text = "" Then
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
